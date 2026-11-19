<#
.SYNOPSIS
    Zoekt alle e-mails van een afzender via Purview Content Search en soft-delete ze na bevestiging.

.DESCRIPTION
    Maakt verbinding met Security & Compliance PowerShell en maakt een Content Search
    (New-ComplianceSearch) aan over alle Exchange-mailboxen met de query 'From:<afzender>'.
    Het script wacht tot de search voltooid is en toont het aantal gevonden items.

    Daarna wordt (na een J/N-bevestiging) een purge-actie van het type SoftDelete gestart
    (New-ComplianceSearchAction -Purge -PurgeType SoftDelete). Soft-deleted items komen in
    de map 'Herstelbare items' van de gebruiker en zijn binnen de bewaartermijn te herstellen.

    Met -WhatIf wordt de search wel uitgevoerd (alleen lezen), maar wordt er niets verwijderd;
    het script toont dan alleen wat er verwijderd zou worden. Met -Force wordt de J/N-vraag
    overgeslagen (bijvoorbeeld voor gebruik in automatisering).

    De search blijft na afloop bestaan in Purview, zodat het resultaat te controleren is.

.PARAMETER SenderAddress
    E-mailadres van de afzender (bijv. afzender@voorbeeld.nl). Wordt gevraagd als je het weglaat.

.PARAMETER PollIntervalSeconds
    Aantal seconden tussen statuscontroles van de search en de purge (standaard 10).

.PARAMETER Force
    Slaat de J/N-bevestiging over en start de soft delete direct.

.EXAMPLE
    .\Remove-EmailFromSender.ps1

    Vraagt om het afzenderadres, voert de search uit en vraagt (J/N) of de mails verwijderd moeten worden.

.EXAMPLE
    .\Remove-EmailFromSender.ps1 -SenderAddress 'afzender@voorbeeld.nl' -WhatIf

    Voert de search uit en laat zien wat er verwijderd zou worden, zonder iets te verwijderen.

.EXAMPLE
    .\Remove-EmailFromSender.ps1 -SenderAddress 'afzender@voorbeeld.nl' -Force

    Voert de search uit en soft-delete de gevonden mails zonder J/N-vraag.

.NOTES
    Auteur   : RJ Lachman
    Versie   : 2.0.0
    Vereisten: Module ExchangeOnlineManagement (Connect-IPPSSession)
    Rollen   : Compliance Search; voor verwijderen ook Search And Purge
    Let op   : Een purge verwijdert maximaal 10 items per mailbox per actie (limiet van Microsoft).
               Voorheen: SoftDelete-EmailFromSender.ps1
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter()]
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$SenderAddress,

    [Parameter()]
    [ValidateRange(5, 300)]
    [int]$PollIntervalSeconds = 10,

    [Parameter()]
    [switch]$Force
)

# -WhatIf onthouden. Verbinden en zoeken zijn alleen-lezen en moeten ook bij -WhatIf werken,
# daarom staat WhatIf tijdelijk uit tot aan de daadwerkelijke verwijderactie.
$whatIfRequested = [bool]$WhatIfPreference
$WhatIfPreference = $false

# --- Verbinden ----------------------------------------------------------------
Write-Host 'Verbinden met Security & Compliance PowerShell...' -ForegroundColor Cyan
try {
    Connect-IPPSSession -EnableSearchOnlySession -ErrorAction Stop
}
catch {
    Write-Host "Verbinden mislukt: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# --- Invoer -------------------------------------------------------------------
if (-not $SenderAddress) {
    $SenderAddress = (Read-Host 'Geef het e-mailadres van de afzender op (bijv. afzender@voorbeeld.nl)').Trim()
}
if ([string]::IsNullOrWhiteSpace($SenderAddress)) {
    Write-Host 'Geen e-mailadres opgegeven. Script stopt.' -ForegroundColor Red
    return
}
if ($SenderAddress -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
    Write-Host "'$SenderAddress' is geen geldig e-mailadres. Script stopt." -ForegroundColor Red
    return
}

# Unieke searchnaam genereren
$searchName = 'DeleteFrom_' + ($SenderAddress -replace '@|\.', '_') + '_' + (Get-Date -Format 'yyyyMMdd_HHmmss')
$query = "From:$SenderAddress"

Write-Host "`nSearch aanmaken: $searchName" -ForegroundColor Yellow
Write-Host "Query: $query" -ForegroundColor Gray

# --- Content Search -----------------------------------------------------------
try {
    $search = New-ComplianceSearch -Name $searchName -ExchangeLocation All -ContentMatchQuery $query `
        -Description "Zoeken en eventueel soft-deleten van mails van $SenderAddress" -ErrorAction Stop
    Start-ComplianceSearch -Identity $search.Identity -ErrorAction Stop
}
catch {
    Write-Host "Search aanmaken of starten mislukt: $($_.Exception.Message)" -ForegroundColor Red
    return
}

Write-Host "`nSearch gestart. Wachten tot deze voltooid is..." -ForegroundColor Cyan

do {
    Start-Sleep -Seconds $PollIntervalSeconds
    try {
        $result = Get-ComplianceSearch -Identity $searchName -ErrorAction Stop
    }
    catch {
        Write-Host "Status ophalen mislukt: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
    $status = $result.Status
    Write-Host "$(Get-Date -Format 'HH:mm:ss') - Status: $status | Items: $($result.Items) | Size: $($result.Size)" -ForegroundColor Gray
    if ($status -eq 'Failed') {
        Write-Host "Search is mislukt. Controleer '$searchName' in de Purview-portal." -ForegroundColor Red
        return
    }
} while ($status -ne 'Completed')

Write-Host "`nSearch voltooid!" -ForegroundColor Green
Write-Host '----------------------------------------'
$result | Format-List Name, Status, Items, Size, ContentMatchQuery, CreatedTime, LastModifiedTime
Write-Host "----------------------------------------`n"

if ($result.Items -eq 0) {
    Write-Host 'Geen mails gevonden. Er valt niets te verwijderen.' -ForegroundColor Yellow
    Write-Host "`nKlaar. Searchnaam: $searchName" -ForegroundColor Cyan
    return
}

# --- Bevestiging --------------------------------------------------------------
$WhatIfPreference = $whatIfRequested
if (-not $whatIfRequested -and -not $Force) {
    $confirm = Read-Host 'Wil je de gevonden mails soft-deleten? (J/N)'
    if ($confirm -notin @('J', 'j', 'Ja', 'ja')) {
        Write-Host "`nSoft delete overgeslagen. De search blijft bestaan in Purview." -ForegroundColor Yellow
        Write-Host "`nKlaar. Searchnaam: $searchName" -ForegroundColor Cyan
        return
    }
}
# De J/N-vraag (of -Force) is de bevestiging; geen tweede ShouldProcess-prompt tenzij -Confirm is opgegeven.
if (-not $PSBoundParameters.ContainsKey('Confirm')) { $ConfirmPreference = 'None' }

# --- Soft delete --------------------------------------------------------------
if ($PSCmdlet.ShouldProcess("$($result.Items) item(s) van $SenderAddress (search '$searchName')", 'Soft delete (Purge)')) {
    Write-Host "`nSoft delete starten..." -ForegroundColor Yellow
    try {
        New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType SoftDelete -Confirm:$false -ErrorAction Stop
    }
    catch {
        Write-Host "Purge starten mislukt: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    Write-Host 'Wachten tot de purge voltooid is...' -ForegroundColor Cyan
    do {
        Start-Sleep -Seconds $PollIntervalSeconds
        # SilentlyContinue: de purge-actie is direct na het starten soms nog niet op te vragen.
        $purge = Get-ComplianceSearchAction -Identity "${searchName}_Purge" -ErrorAction SilentlyContinue
        if ($purge) {
            Write-Host "$(Get-Date -Format 'HH:mm:ss') - Purge status: $($purge.Status)" -ForegroundColor Gray
            if ($purge.Status -eq 'Failed') {
                Write-Host 'Purge is mislukt.' -ForegroundColor Red
                $purge | Format-List SearchName, Status, Errors, Results
                return
            }
        }
        else {
            Write-Host "$(Get-Date -Format 'HH:mm:ss') - Purge actie nog niet beschikbaar..." -ForegroundColor DarkGray
        }
    } while ($purge.Status -ne 'Completed')

    Write-Host "`nSoft delete voltooid!" -ForegroundColor Green
    $purge | Format-List SearchName, Status, JobStartTime, JobEndTime, Results
}

Write-Host "`nKlaar. Searchnaam: $searchName" -ForegroundColor Cyan
