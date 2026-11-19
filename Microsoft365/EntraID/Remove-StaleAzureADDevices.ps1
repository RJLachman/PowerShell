<#
.SYNOPSIS
    Exporteert en verwijdert (na bevestiging) Entra ID-apparaten die al X dagen niet zijn aangemeld.

.DESCRIPTION
    Gebruikt Microsoft Graph PowerShell (vervanger van de uitgefaseerde AzureAD-module).
    - Haalt alle apparaten op met een ApproximateLastSignInDateTime ouder dan de drempel.
      Apparaten die nog nooit zijn aangemeld (geen datum) worden overgeslagen.
    - Exporteert ze naar een CSV (standaard EntraDeviceExport.csv naast het script).
    - Vraagt bevestiging (J/N) en verwijdert ze daarna.

    De Graph-login gebruikt geen WAM (Windows Web Account Manager), omdat die in o.a. PowerShell ISE
    en embedded terminals faalt ("A window handle must be configured"). Mislukt de interactieve
    login, dan valt het script terug op device code login.

    Met -WhatIf wordt de export wel gemaakt, maar wordt er niets verwijderd.
    Met -Force wordt de J/N-vraag overgeslagen.

.PARAMETER DeletionThresholdDays
    Aantal dagen zonder aanmelding waarna een apparaat wordt verwijderd (standaard 60).

.PARAMETER ExportPath
    Pad van het CSV-exportbestand (standaard EntraDeviceExport.csv in de map van het script).

.PARAMETER Force
    Slaat de J/N-bevestiging over en verwijdert direct.

.EXAMPLE
    .\Remove-StaleAzureADDevices.ps1

    Exporteert apparaten die langer dan 60 dagen niet zijn aangemeld en vraagt (J/N) om ze te verwijderen.

.EXAMPLE
    .\Remove-StaleAzureADDevices.ps1 -DeletionThresholdDays 90 -WhatIf

    Toont welke apparaten (ouder dan 90 dagen) verwijderd zouden worden, zonder iets te verwijderen.

.EXAMPLE
    .\Remove-StaleAzureADDevices.ps1 -DeletionThresholdDays 120 -ExportPath 'C:\temp\stale.csv' -Force

    Exporteert naar C:\temp\stale.csv en verwijdert zonder J/N-vraag.

.NOTES
    Auteur   : RJ Lachman
    Versie   : 2.1.0
    Vereisten: Install-Module Microsoft.Graph.Identity.DirectoryManagement
    Rechten  : Device.ReadWrite.All (Cloud Device Administrator of hoger)
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter()]
    [ValidateRange(1, 3650)]
    [int]$DeletionThresholdDays = 60,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ExportPath = (Join-Path -Path $PSScriptRoot -ChildPath 'EntraDeviceExport.csv'),

    [Parameter()]
    [switch]$Force
)

# Verbinden, ophalen en exporteren zijn ook bij -WhatIf gewenst; WhatIf geldt alleen voor het verwijderen.
$whatIfRequested = [bool]$WhatIfPreference
$WhatIfPreference = $false

# --- Verbinden ----------------------------------------------------------------
try {
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
}
catch {
    Write-Host "Module Microsoft.Graph.Identity.DirectoryManagement ontbreekt: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'Installeer met: Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser' -ForegroundColor Yellow
    return
}

if (Get-Command Set-MgGraphOption -ErrorAction SilentlyContinue) {
    Set-MgGraphOption -DisableLoginByWAM $true
}
try {
    Connect-MgGraph -Scopes 'Device.ReadWrite.All' -NoWelcome -ErrorAction Stop
}
catch {
    Write-Warning "Interactieve login mislukt ($($_.Exception.Message)). Terugvallen op device code login..."
    try {
        Connect-MgGraph -Scopes 'Device.ReadWrite.All' -UseDeviceCode -NoWelcome -ErrorAction Stop
    }
    catch {
        Write-Host "Verbinden met Microsoft Graph mislukt: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}

# --- Apparaten ophalen --------------------------------------------------------
$deletionThreshold = (Get-Date).AddDays(-$DeletionThresholdDays)

Write-Host 'Apparaten ophalen...' -ForegroundColor Cyan
try {
    $staleDevices = @(
        Get-MgDevice -All -ErrorAction Stop -Property Id, DeviceId, DisplayName, ApproximateLastSignInDateTime,
        OperatingSystem, OperatingSystemVersion, IsCompliant, IsManaged |
            Where-Object { $_.ApproximateLastSignInDateTime -and $_.ApproximateLastSignInDateTime -le $deletionThreshold }
    )
}
catch {
    Write-Host "Apparaten ophalen mislukt: $($_.Exception.Message)" -ForegroundColor Red
    return
}

if ($staleDevices.Count -eq 0) {
    Write-Host "Geen apparaten gevonden die langer dan $DeletionThresholdDays dagen niet zijn ingelogd." -ForegroundColor Green
    return
}

# --- Exporteren ---------------------------------------------------------------
try {
    $staleDevices |
        Select-Object DisplayName, Id, DeviceId, ApproximateLastSignInDateTime, OperatingSystem, OperatingSystemVersion, IsCompliant, IsManaged |
        Export-Csv -Path $ExportPath -UseCulture -NoTypeInformation -ErrorAction Stop
}
catch {
    Write-Host "Exporteren naar $ExportPath mislukt: $($_.Exception.Message)" -ForegroundColor Red
    return
}

Write-Host "$($staleDevices.Count) apparaten gevonden. Rapport: $ExportPath" -ForegroundColor Yellow

# --- Bevestiging --------------------------------------------------------------
$WhatIfPreference = $whatIfRequested
if (-not $whatIfRequested -and -not $Force) {
    $confirmDeletion = $null
    while ($confirmDeletion -notin @('J', 'j', 'Ja', 'ja', 'N', 'n', 'Nee', 'nee')) {
        $confirmDeletion = Read-Host "Alle apparaten verwijderen die sinds $DeletionThresholdDays dagen niet zijn ingelogd? (J/N)"
    }
    if ($confirmDeletion -notin @('J', 'j', 'Ja', 'ja')) {
        Write-Host 'Geannuleerd. Er is niets verwijderd.' -ForegroundColor Yellow
        return
    }
}
# De J/N-vraag (of -Force) is de bevestiging; geen tweede prompt per apparaat tenzij -Confirm is opgegeven.
if (-not $PSBoundParameters.ContainsKey('Confirm')) { $ConfirmPreference = 'None' }

# --- Verwijderen --------------------------------------------------------------
$success = 0
$failed = 0
foreach ($device in $staleDevices) {
    $target = "$($device.DisplayName) ($($device.Id)), laatst aangemeld $($device.ApproximateLastSignInDateTime)"
    if (-not $PSCmdlet.ShouldProcess($target, 'Apparaat verwijderen uit Entra ID')) { continue }
    try {
        Remove-MgDevice -DeviceId $device.Id -ErrorAction Stop
        Write-Host "Verwijderd: $($device.DisplayName) ($($device.Id))" -ForegroundColor Green
        $success++
    }
    catch {
        Write-Host "Mislukt: $($device.DisplayName) - $($_.Exception.Message)" -ForegroundColor Red
        $failed++
    }
}

if (-not $whatIfRequested) {
    Write-Host "`nKlaar. Succesvol: $success | Mislukt: $failed" -ForegroundColor Cyan
}
