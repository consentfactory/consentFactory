$users = Import-Csv -Path C:\lyncScripts\csv\CommonAreaPhones.csv
foreach ($user in $users) {
    $displayName = $($user.Name)
    $ext = $($user.Ext)
    $telFour = $($user.Phone).Substring(6)
    $location = $($user.Location)
    $sip = "sip:$($user.SIP)"

    #Get Voice Policy based on location info
    if ($location -like "*District*") {$voicePolicy = "DistrictOffice-General"} 
    elseif ($location -like "*Special*") {$voicePolicy = "SpecialServices-General"}
    elseif ($location -like "*Food*") {$voicePolicy = "FoodServices-General"}
    elseif ($location -like "*Jefferson*") {$voicePolicy = "JeffersonES-General"}
    
    <#Write-Host "$($displayName) - $($location)"
    Write-Host "Phone: tel:+1555667$($telFour);ext=$($ext)"
    Write-Host $voicePolicy#>
    Set-CsCommonAreaPhone `
    -Identity $displayName `
    -SipAddress $sip `
    -Verbose
    Set-CsClientPin -Identity "$displayName" -Pin 123123 -Verbose
    Get-CsCommonAreaPhone $displayName | Grant-CsVoicePolicy -PolicyName $voicePolicy -Verbose
}