# SoftDelete-EmailFromSender.ps1
# Script om mails van een specifiek afzender te zoeken en optioneel soft te deleten

# Verbinden met Security & Compliance PowerShell
Write-Host "Verbinden met Security & Compliance PowerShell..." -ForegroundColor Cyan
Connect-IPPSSession -EnableSearchOnlySession

# Emailadres opvragen
$sender = Read-Host "Geef het e-mailadres van de afzender op (bijv. afzender@voorbeeld.nl)"

if ([string]::IsNullOrWhiteSpace($sender)) {
    Write-Host "Geen e-mailadres opgegeven. Script stopt." -ForegroundColor Red
    return
}

# Unieke searchnaam genereren
$searchName = "DeleteFrom_" + ($sender -replace '@|\.', '_') + "_" + (Get-Date -Format "yyyyMMdd_HHmmss")

Write-Host "`nSearch aanmaken: $searchName" -ForegroundColor Yellow
Write-Host "Query: From:$sender" -ForegroundColor Gray

# Content Search aanmaken en starten
$search = New-ComplianceSearch -Name $searchName -ExchangeLocation All -ContentMatchQuery "From:$sender" -Description "Zoeken en eventueel soft-deleten van mails van $sender"
Start-ComplianceSearch -Identity $search.Identity

Write-Host "`nSearch gestart. Wachten tot deze voltooid is..." -ForegroundColor Cyan

# Loopje om te wachten op Completed
do {
    Start-Sleep -Seconds 10
    $result = Get-ComplianceSearch -Identity $searchName
    $status = $result.Status
    $items = $result.Items
    $size = $result.Size
    Write-Host "$(Get-Date -Format 'HH:mm:ss') - Status: $status | Items: $items | Size: $size" -ForegroundColor Gray
} while ($status -ne "Completed")

Write-Host "`nSearch voltooid!" -ForegroundColor Green
Write-Host "----------------------------------------"
$result | Format-List Name, Status, Items, Size, ContentMatchQuery, CreatedTime, LastModifiedTime
Write-Host "----------------------------------------`n"

# Vragen of soft delete gewenst is
$confirm = Read-Host "Wil je de gevonden mails soft-deleten? (J/N)"

if ($confirm -eq "J" -or $confirm -eq "j" -or $confirm -eq "Ja" -or $confirm -eq "ja") {
    Write-Host "`nSoft delete starten..." -ForegroundColor Yellow
    New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType SoftDelete -Confirm:$false

    Write-Host "Wachten tot de purge voltooid is..." -ForegroundColor Cyan
    do {
        Start-Sleep -Seconds 10
        $purge = Get-ComplianceSearchAction -Identity "${searchName}_Purge" -ErrorAction SilentlyContinue
        if ($purge) {
            Write-Host "$(Get-Date -Format 'HH:mm:ss') - Purge status: $($purge.Status)" -ForegroundColor Gray
        } else {
            Write-Host "$(Get-Date -Format 'HH:mm:ss') - Purge actie nog niet beschikbaar..." -ForegroundColor DarkGray
        }
    } while ($purge.Status -ne "Completed")

    Write-Host "`nSoft delete voltooid!" -ForegroundColor Green
    $purge | Format-List SearchName, Status, JobStartTime, JobEndTime, Results
} else {
    Write-Host "`nSoft delete overgeslagen. De search blijft bestaan in Purview." -ForegroundColor Yellow
}

Write-Host "`nKlaar. Searchnaam: $searchName" -ForegroundColor Cyan
