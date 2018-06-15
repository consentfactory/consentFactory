<#
Functions
Written 20161115
Contains set of functions used for deploying Skype for Business accounts
Exchange as part of unified messaging
#>

Function New-SkypeUser {
    #$tel is probably a bit a redundant, but I left it in for the flexibility later if needed. However, we could delete this,
    #we just need to change $phone to be the four-digit number after the NPA-NXX, and change the commands below
    $tel = $phone.Substring(6)
    $regPool = "pool.contoso.org"
    $dc = "dc01.contoso.org"

    #Getting UPN from AD so that we can update the telephone number later
    $ADUser = Get-ADUser -filter {UserPrincipalName -eq $user}

    #Enable the AD User for Skype for Business
    $userExists = Get-CSUser $user -ErrorAction SilentlyContinue
    if ($userExists) {
        Write-Host "User $($user) is already enabled. Modifying..." -ForegroundColor Green
    } else {
        Try {
            Write-Host "Enabling $($user) for Skype for Business..." -ForegroundColor Green
            Enable-CSuser $user -RegistrarPool $regPool -SipAddressType UserPrincipalName -WarningAction SilentlyContinue

            #Pause script for 45 seconds to allow updates, then send a AD Sites and Services replication command, otherwise updating PIN will fail
            Write-Host "Replicating user changes to domain controller. Please wait 40 seconds..." -ForegroundColor Green
            Start-Sleep -s 15 
            Invoke-Command -ComputerName $dc -ScriptBlock {repadmin /syncall} | Out-Null
            Invoke-Command -ComputerName $dc -ScriptBlock {repadmin /syncall} | Out-Null
            Invoke-Command -ComputerName $dc -ScriptBlock {repadmin /syncall} | Out-Null
            Write-Host "15 more seconds..." -ForegroundColor Green
            Start-Sleep -s 15
        } Catch {
            $_.Exception.Message
            break
            }
    }

    #Enable user for enterprise voice
    Try {
        Write-Host "Enabling\modifying $($user) for enterprise voice..." -ForegroundColor Green     
        Set-CSUser $user -EnterpriseVoiceEnabled $true -LineUri "tel:+1555667$($tel);ext=$($ext)"  -ErrorAction Stop -WarningAction SilentlyContinue
    } Catch {
        $errorMsg = $_.Exception.Message
        if ($errorMsg -like "*Filter failed to return unique result*") {
        Write-Host "Extension is already in use by:" -ForegroundColor Magenta
        Get-CSUSer | where {$_.lineuri -like "*$($ext)*"} | select displayName,lineuri | fl
        break
        }
    }

    #Grant the appropriate voice policy
    Write-Host "Granting $($user) voice policy "$($voicePolicy)"..." -ForegroundColor Green
    Get-CSUser $user | Grant-CsVoicePolicy -PolicyName $voicePolicy -WarningAction SilentlyContinue

    #Set PIN for PIN authentication for phones.
    Write-Host "Setting up PIN for phone login..." -ForegroundColor Green
    Try {        
        Get-CSUser $user | Set-CsClientPin -Pin 123123 -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null
    } Catch {
        $errorMsg = $_.Exception.Message
        if ($errorMsg -like "*is not validly homed on target*") {
        Write-Host "User information did not update in time for PIN to be set. Please run again when script is done." -ForegroundColor Magenta
        }
    } 

    #Update AD telephone number property
    Write-Host "Updating telephone number in Active Directory..." -ForegroundColor Green
    Set-ADUser $ADUser.SamAccountName -OfficePhone "+1 (555) 667-$($tel) x$($ext)"
    Write-Host "`n"
}

Function enable-skypeVoiceMail {
    #Getting UPN from AD
    $ADUser = Get-ADUser -filter {UserPrincipalName -eq $user}

    #Establish session with Exchange Server
    Write-Host "Establishing connection with Exchange server..." -ForegroundColor Yellow
    $username = "skypeExchangeSVC"
    $password = Get-Content -Path C:\skypeScripts\key.txt
    $creds = Get-Credential
    #$creds = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $username,($password | ConvertTo-SecureString)
    $timeout = New-PSSessionOption -IdleTimeout 120000
    $Session = New-PSSession `
    -ConfigurationName Microsoft.Exchange `
    -ConnectionUri http://exchange.contoso.org/PowerShell/ `
    -Credential $creds `
    -Authentication Kerberos `
    -AllowRedirection `
    -SessionOption $timeout `
    -WarningAction SilentlyContinue
    Try {
        Import-PSSession $Session -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null
    } Catch {
        $errorMsg = $_.Exception.Message
        if ($errorMsg -like "*No command proxies have been created*") {
            $errorMsg | Out-Null
        } else {
            $errorMsg
            Break
        }
    }
    Write-Host "Beginning setting up voicemail..." -ForegroundColor Yellow
    Write-Host "UPN: "$user -ForegroundColor Yellow
    Write-Host "Extension: "$ext -ForegroundColor Yellow
    
    #Create mailbox contact that we will forward email to
    try {
        $mailContactCheck = Get-MailContact "$($user)-GMAIL" -ErrorAction Stop
    } catch {
        $errorMsg = $_.Exception.Message
        if ($errorMsg -notlike "*couldn't be found on*") {
            $errorMsg
            break
        } else {
            $errorMsg | Out-Null
        }
    }
    if (!$mailContactCheck) {
        Write-Host "Creating Mailbox Contact $($user)..." -ForegroundColor Yellow
        New-MailContact -name "$($user)-GMAIL" -ExternalEmailAddress $user -OrganizationalUnit "OU=VM Contacts,OU=Skype for Business,DC=contoso,DC=org" | Out-Null
    }

    #Enable mailbox for the user, and grant a retention policy that removes voicemails and emails after 30 days
    try {
        $mailboxCheck = Get-Mailbox $user -ErrorAction Stop
    } catch {
        $errorMsg = $_.Exception.Message
        if ($errorMsg -notlike "*couldn't be found on*") {
            $errorMsg
            break
        } else {
            $errorMsg | Out-Null
        }
    }
    if (!$mailboxCheck) {
        Write-Host "Enabling Mailbox..." -ForegroundColor Yellow
        Enable-Mailbox $user -Database mailboxes01 -RetentionPolicy "Delete Voicemails After 30 Days" | Out-Null

        ##Pause script for 30 seconds to allow updates, then send a AD Sites and Services replication command,
        ##otherwise the server/DC doesn't have time to see the new-enabled mailbox, and the UM update will fail.
        Write-Host "Replicating user changes to domain controller. Please wait 40 seconds" -ForegroundColor Yellow
        Start-Sleep -s 15 
        Invoke-Command -ComputerName $dc -ScriptBlock {repadmin /syncall} | Out-Null
        Invoke-Command -ComputerName $dc -ScriptBlock {repadmin /syncall} | Out-Null
        Invoke-Command -ComputerName $dc -ScriptBlock {repadmin /syncall} | Out-Null
        Write-Host "15 more seconds..." -ForegroundColor Yellow
        Start-Sleep -s 15

        #Getting email address. Deprecated item since my first version of this script. Don't really need this or the Write-Host following this.
        #Write-Host "Getting email address for a variable..." -ForegroundColor Yellow
        $email = Get-Mailbox $user | Sort-Object PrimarySmtpAddress | Select-Object -ExpandProperty PrimarySmtpAddress
        #Write-Host "Email: "$email -ForegroundColor Green

        #Set mailbox to forward emails to the corresponding Gmail account, created above
        Write-Host "Forwarding email to Gmail..." -ForegroundColor Yellow
        Set-Mailbox $user -DeliverToMailboxAndForward $true -ForwardingAddress "$($user)-GMAIL" -WarningAction SilentlyContinue

        #Enabled the mailbox for UM
        Write-Host "Enabling Unified Messaging for user $($user)..." -ForegroundColor Yellow
        Enable-UMMailbox $user -SIPResourceIdentifier $user -Extensions $ext -UMMailboxPolicy "JeromeSD Default Policy" -PIN $voicemailPIN -PinExpired $true | Out-Null
    
        #These settings prevent the users internally from adjusting any settings regarding their mailbox, or link up to the phone.
        #However, these users may still be able to use Outlook, which couldn't be deactivated because S4B uses the same protocols for access
        Set-CASMailbox $user -ImapEnabled $false -PopEnabled $false -OWAEnabled $false -ActiveSyncEnabled $false -WarningVariable $wv | Out-Null
        Write-Host "`n"
    } else {

        #Update extension for user
        Write-Host "Updating extension in Unified Messaging..." -ForegroundColor Yellow
        $mailbox = Get-Mailbox $user
        [string]$oldUMExt = $mailbox.EmailAddresses | Select-String -Pattern '\d{4}' | ForEach-Object {$_.Matches[0].Value} | Out-String
        [string]$oldUMExtURI = "eum:$($oldUMExt);phone-context=contoso1.contoso.org" | Out-String
        [string]$newUMExtURI = "eum:$($ext);phone-context=contoso1.contoso.org" | Out-String
        Set-ADUser $ADUser.SamAccountName -Remove @{proxyAddresses=$oldUMExtURI}
        Set-ADUser $ADUser.SamAccountName -Add @{proxyAddresses=$newUMExtURI}
        <#$mailbox.EmailAddresses = $mailbox.EmailAddresses | Where-Object { $_ } | Select -Unique
        $mailbox.emailaddresses.Remove($oldUMExtURI)
        $mailbox.EmailAddresses = $mailbox.EmailAddresses | Where-Object { $_ } | Select -Unique
        $mailbox.emailaddresses.Add($newUMExtURI)#>
        Try {
            Set-Mailbox $user -EmailAddresses $mailbox.emailaddresses -ErrorAction Stop
        } Catch {
            $errorMsg = $_.Exception.Message
            if ($errorMsg -notlike "*is already present in the collection.*") {
                $errorMsg
                break
            } else {
                Write-Host "No changes made. Extension is already configured." -ForegroundColor Yellow
            }
        }
    }
    Write-Host "`n"
}

function userCompleted {
$completed = Get-CSUSer $user | Select sipaddress,lineuri,voicepolicy
Write-Host "SIP Address: "$completed.sipaddress
Write-Host "Phone and Extension: "$completed.lineuri
Write-Host "Voice Policy: "$completed.voicepolicy
}