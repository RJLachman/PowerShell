<#
.SYNOPSIS
 Exporteert en verwijdert (na bevestiging) Entra ID-apparaten die al X dagen niet zijn ingelogd.

.DESCRIPTION
 Gebruikt Microsoft Graph PowerShell (vervanger van de uitgefaseerde AzureAD-module).
 - Haalt alle apparaten op met ApproximateLastSignInDateTime ouder dan de drempel.
 - Exporteert ze naar een CSV naast het script.
 - Vraagt bevestiging (J/N) en verwijdert ze daarna.

.PARAMETER DeletionThresholdDays
 Aantal dagen zonder aanmelding (standaard 60).

.NOTES
 Vereisten: Install-Module Microsoft.Graph.Identity.DirectoryManagement
 Rechten : Device.ReadWrite.All (Cloud Device Administrator of hoger)
#>

[CmdletBinding()]
param(
    [int]$DeletionThresholdDays = 60
)

Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
# WAM-login (Windows Web Account Manager) faalt in o.a. PowerShell ISE en embedded terminals
# ("A window handle must be configured"). Daarom uitzetten en zo nodig terugvallen op device code.
if (Get-Command Set-MgGraphOption -ErrorAction SilentlyContinue) {
    Set-MgGraphOption -DisableLoginByWAM $true
}
try {
    Connect-MgGraph -Scopes "Device.ReadWrite.All" -NoWelcome -ErrorAction Stop
} catch {
    Write-Warning "Interactieve login mislukt ($($_.Exception.Message)). Terugvallen op device code login..."
    Connect-MgGraph -Scopes "Device.ReadWrite.All" -UseDeviceCode -NoWelcome -ErrorAction Stop
}

$deletionThreshold = (Get-Date).AddDays(-$DeletionThresholdDays)

Write-Host "Apparaten ophalen..." -ForegroundColor Cyan
$allDevices = Get-MgDevice -All -Property Id, DeviceId, DisplayName, ApproximateLastSignInDateTime, OperatingSystem, OperatingSystemVersion, IsCompliant, IsManaged |
    Where-Object { $_.ApproximateLastSignInDateTime -and $_.ApproximateLastSignInDateTime -le $deletionThreshold }

if (-not $allDevices) {
    Write-Host "Geen apparaten gevonden die langer dan $DeletionThresholdDays dagen niet zijn ingelogd." -ForegroundColor Green
    return
}

$exportPath = Join-Path $PSScriptRoot "EntraDeviceExport.csv"
$allDevices |
    Select-Object DisplayName, Id, DeviceId, ApproximateLastSignInDateTime, OperatingSystem, OperatingSystemVersion, IsCompliant, IsManaged |
    Export-Csv -Path $exportPath -UseCulture -NoTypeInformation

Write-Host "$(@($allDevices).Count) apparaten gevonden. Rapport: $exportPath" -ForegroundColor Yellow

$confirmDeletion = $null
while ($confirmDeletion -notin @('J','j','Ja','ja','N','n','Nee','nee')) {
    $confirmDeletion = Read-Host "Alle apparaten verwijderen die sinds $DeletionThresholdDays dagen niet zijn ingelogd? (J/N)"
}

if ($confirmDeletion -in @('J','j','Ja','ja')) {
    $success = 0; $failed = 0
    foreach ($device in $allDevices) {
        try {
            Remove-MgDevice -DeviceId $device.Id -ErrorAction Stop
            Write-Host "Verwijderd: $($device.DisplayName) ($($device.Id))" -ForegroundColor Green
            $success++
        } catch {
            Write-Host "Mislukt: $($device.DisplayName) - $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        }
    }
    Write-Host "`nKlaar. Succesvol: $success | Mislukt: $failed" -ForegroundColor Cyan
} else {
    Write-Host "Geannuleerd. Er is niets verwijderd." -ForegroundColor Yellow
}
