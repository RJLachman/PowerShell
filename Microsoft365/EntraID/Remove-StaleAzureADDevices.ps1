#enter treshold days
$deletionTresholdDays= 60

Connect-AzureAD -ErrorAction Stop

$deletionTreshold= (Get-Date).AddDays(-$deletionTresholdDays)

$allDevices=Get-AzureADDevice -All:$true | Where {$_.ApproximateLastLogonTimeStamp -le $deletionTreshold}

$exportPath=$(Join-Path $PSScriptRoot "AzureADDeviceExport.csv")

$allDevices | Select-Object -Property DisplayName, ObjectId, ApproximateLastLogonTimeStamp, DeviceOSType, DeviceOSVersion, IsCompliant, IsManaged `
| Export-Csv -Path $exportPath -UseCulture -NoTypeInformation

Write-Output "Find report with all devices under: $exportPath"

$confirmDeletion=$null

while ($confirmDeletion -notin @('J','j','Ja','ja','N','n','Nee','nee')){

    $confirmDeletion = Read-Host "Delete all Azure AD devices which haven't contacted your tenant since $deletionTresholdDays days (J/N)"
}

if ($confirmDeletion -in @('J','j','Ja','ja')){

    $allDevices | ForEach-Object {

        Write-Output "Removing device $($PSItem.ObjectId)"

        Remove-AzureADDevice -ObjectId $PSItem.ObjectId
    }

} else {
  
    Write-Output "Exiting..."
}
