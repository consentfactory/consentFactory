<# 
newUser.ps1 - Version 2.0 - Updated: 11/14/16 - Created by Jimmy Taylor
-This script creates new individual users for Skype for Business.
-You must know UPN/SIP address, phone number, extension, and voice policy.
-This script assumes UPN and SIP address will be the same

UPDATES
2.0
-Rewrote most of the script
-Moved main commands to function.ps1
-Added Exchange features so that user creation and changes were combined
-Script now creates and updates users in single process
-Gives information about changes or errors that are more descriptive

1.3
-First version with documentation.
-Added comments.
-Basic script to create new users in Skype for Business.
#>

#*** ENTER THE INFORMATION BELOW ***


#UPN, phone, and extension variables
$Global:user = "test.user9@contoso.org"
$Global:phone = "5556672392"
$Global:ext = "9993"
$Global:voicemailPIN = "991456"

<# Choose your voice policy
DistrictOffice-General
FoodService-General
SpecialServices-General
#>

#Voice policy name that we're going to grant to the user; copy/paste from above
$Global:voicePolicy = "DistrictOffice-General"

#
#
#
#
#
# ***** Don't adjust anything below this! *****
#

#region functions
#Importing commands from separate file
. 'C:\skypeScripts\functions.ps1'

Write-Host "`n"
#Enable user for Skype for Business
new-SkypeUser
#Enable user for Exchange Voicemail
enable-skypeVoiceMail

Write-Host "Done with user $($user)!" -ForegroundColor Green
userCompleted
#endregion
