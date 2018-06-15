<#
enableMail_FWD2GMAIL.ps1
Created 20160915

The purpose of this script is to create mailbox contacts for
corresponding Gmail accounts. This works as a sort of half measure of getting
some of the benefits of Exchange UM with Skype for Business when
operating as a Google shop.

The major benefit of this is that missed calls, voicemails, and other notifications are
sent to Gmail, but they can't be managed via Gmail, which is why this is kind of a half-UM
method. 

Script assumes that internal mailboxes are of a different domain (e.g., "user@contoso.internal")
and that mail flows have been properly configured
#>

$users = Import-Csv -Path C:\scripts\csv\userUpdate.csv

foreach ($user in $users) {
    $upn = $($user.EmailAddress)
    $ext = $($user.ext)
    $pin = $($user.pin)

    Write-Host "UPN: "$upn -ForegroundColor Green
    Write-Host "Extension: "$ext -ForegroundColor Green
    Write-Host "Creating Mailbox Contact $($upn)..." -ForegroundColor Yellow

    # Create Gmail mailbox contact
    New-MailContact -name "$($upn)-GMAIL" -ExternalEmailAddress $upn -OrganizationalUnit "OU=VM Contacts,OU=Skype for Business,DC=contoso,DC=org"
    Write-Host "`n"
    Write-Host "Enabling Mailbox..." -ForegroundColor Yellow
    
    # Enabling mailbox. Note: UPN will not be the same as internal email address
    # Retention policy is also configured for mailbox that deletes voicemails after 30 days
    Enable-Mailbox $upn -Database mailboxes01 -RetentionPolicy "Delete Voicemails After 30 Days"
    Write-Host "`n"
    Write-Host "Getting email address for a variable..." -ForegroundColor Yellow

    $email = Get-Mailbox $upn | Sort-Object PrimarySmtpAddress | Select-Object -ExpandProperty PrimarySmtpAddress
    Write-Host "Email: "$email -ForegroundColor Green
    Write-Host "`n"
    Write-Host "Forwarding email to Gmail..." -ForegroundColor Yellow
    
    # Setting up internal email account (@contoso.internal) to forward to Gmail account
    Set-Mailbox $upn -DeliverToMailboxAndForward $true -ForwardingAddress "$($upn)-GMAIL" -Verbose
    Write-Host "`n"
    Start-Sleep -s 30

    # After setting up mail forward, UM will be enabled, resulting in voicemail instructions
    # and other info to be sent to Gmail account
    Write-Host "Enabling Unified Messaging for user $($upn)..." -ForegroundColor Yellow
    Enable-UMMailbox $upn -SIPResourceIdentifier $upn -Extensions $ext -UMMailboxPolicy "Contoso Default Policy" -PIN $pin -PinExpired $true
    Write-Host "`n"
    
    # Setting up mailbox to not be accessible, except through UM 
    # This is largely a pass-through mailbox, and we want users to manage messages via S4B
    Set-CASMailbox $upn -ImapEnabled $false -PopEnabled $false -OWAEnabled $false -ActiveSyncEnabled $false -MAPIBlockOutlookRpcHttp $true
    Write-Host "Done with user $($upn)!"
}