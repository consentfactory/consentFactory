<# 
newUsers-Bulk.ps1 - Version 2.0 - 11/14/16 - Created by Jimmy Taylor, Tek-Hut
-This script creates users for Skype for Business in bulk.
-The CSV File must have UPN, phone number, extension, and location.
-This script assumes UPN and SIP address will be the same

UPDATES
2.0
-Rewrote most of the script
-Moved main commands to function.ps1
-Added Exchange features so that user creation and changes were combined
-Script now creates and updates users in single process
-Gives information about changes or errors that are more descriptive
-Added prompt so that user can verify filename

1.4
-First version with documentation.
-Added comments.
-Basic script to create new users in Skype for Business.

#>

#***UPDATE PATH TO CSV FILE HERE***
$path = "C:\skypeScripts\csv\userUpdate.csv"

#
# ***** Don't adjust anything below this! *****
#

$accounts = Import-Csv -Path $path
#Importing commands from separate file
. 'C:\skypeScripts\functions.ps1'
Write-Host "`n"

#Set up confimration before running script. Allows user to verify filename.
$confirmation = $null
while ($confirmation -notmatch "[y|n]") {
    $confirmation = Read-Host "Are you sure you want to use the file '$($path.Substring(20))'? (y/n)"
}
#If they say yes, proceed with script, otherwise quit.
if ($confirmation -eq "y") {
    foreach ($account in $accounts) {
        #UPN, phone, and extension variables
        $Script:user = $account.EmailAddress
        $Script:phone = $account.Phone
        $Script:ext = $account.Ext
        $Script:voicemailPIN = $account.Pin
        $Script:voicePolicy = $null

        #Get Voice Policy based on location info from CSV
        if ($location -like "*District*" -or $location -eq $null) {$voicePolicy = "DistrictOffice-General"} 
        elseif ($location -like "*Special*") {$voicePolicy = "SpecialServices-General"}
        elseif ($location -like "*Food*") {$voicePolicy = "FoodServices-General"}
        elseif ($location -like "*Jefferson*") {$voicePolicy = "JeffersonES-General"}

        #region functions

        Write-Host "`n"
        #Enable user for Skype for Business
        new-SkypeUser
        #Enable user for Exchange Voicemail
        enable-skypeVoiceMail

        Write-Host "Done with user $($user)!" -ForegroundColor Green
        userCompleted
        #endregion
    }
} elseif ($confirmation -eq "n") {Exit}