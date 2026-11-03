# Copy-ADGroupMembership.ps1
# Kopieert de group memberships van een brongebruiker naar een doelgebruiker in on-premises AD
# Werkt met e-mailadres

Import-Module ActiveDirectory -ErrorAction Stop

# Bron- en doelgebruiker opvragen (e-mailadres)
$sourceEmail = Read-Host "Geef het e-mailadres van de brongebruiker op"
$targetEmail = Read-Host "Geef het e-mailadres van de doelgebruiker op"

if ([string]::IsNullOrWhiteSpace($sourceEmail) -or [string]::IsNullOrWhiteSpace($targetEmail)) {
    Write-Host "Bron of doel is leeg. Script stopt." -ForegroundColor Red
    return
}

if ($sourceEmail -eq $targetEmail) {
    Write-Host "Bron en doel zijn hetzelfde. Script stopt." -ForegroundColor Red
    return
}

# Gebruikers ophalen op e-mailadres
Write-Host "`nGebruikers ophalen..." -ForegroundColor Yellow
try {
    $source = Get-ADUser -Filter "mail -eq '$sourceEmail'" -Properties MemberOf, mail -ErrorAction Stop
    $target = Get-ADUser -Filter "mail -eq '$targetEmail'" -Properties MemberOf, mail -ErrorAction Stop
} catch {
    Write-Host "Fout bij ophalen gebruiker: $($_.Exception.Message)" -ForegroundColor Red
    return
}

if (-not $source) {
    Write-Host "Brongebruiker met e-mail '$sourceEmail' niet gevonden." -ForegroundColor Red
    return
}
if (-not $target) {
    Write-Host "Doelgebruiker met e-mail '$targetEmail' niet gevonden." -ForegroundColor Red
    return
}
if ($source.Count -gt 1) {
    Write-Host "Meerdere gebruikers gevonden met e-mail '$sourceEmail'. Script stopt." -ForegroundColor Red
    return
}
if ($target.Count -gt 1) {
    Write-Host "Meerdere gebruikers gevonden met e-mail '$targetEmail'. Script stopt." -ForegroundColor Red
    return
}

Write-Host "Bron: $($source.Name) <$($source.mail)>" -ForegroundColor Green
Write-Host "Doel: $($target.Name) <$($target.mail)>" -ForegroundColor Green

# Groepen van de bron ophalen
$sourceGroups = $source.MemberOf

if (-not $sourceGroups) {
    Write-Host "`nDe brongebruiker zit in geen enkele groep." -ForegroundColor Yellow
    return
}

# Groepen van het doel ophalen
$targetGroups = $target.MemberOf

# Groepen die nog niet bij het doel zitten
$groupsToAdd = $sourceGroups | Where-Object { $_ -notin $targetGroups }

Write-Host "`nDe volgende groepen worden gekopieerd ($($groupsToAdd.Count) stuks):" -ForegroundColor Cyan
Write-Host "----------------------------------------"
foreach ($groupDN in $groupsToAdd) {
    $group = Get-ADGroup -Identity $groupDN
    Write-Host "- $($group.Name)" -ForegroundColor White
}
Write-Host "----------------------------------------`n"

if ($groupsToAdd.Count -eq 0) {
    Write-Host "De doelgebruiker zit al in alle groepen van de bron." -ForegroundColor Yellow
    return
}

# Bevestiging vragen
$confirm = Read-Host "Weet je het zeker? Deze memberships kopiëren? (J/N)"

if ($confirm -eq "J" -or $confirm -eq "j" -or $confirm -eq "Ja" -or $confirm -eq "ja") {
    Write-Host "`nKopiëren starten..." -ForegroundColor Yellow
    $success = 0
    $failed = 0

    foreach ($groupDN in $groupsToAdd) {
        try {
            $group = Get-ADGroup -Identity $groupDN
            Add-ADGroupMember -Identity $group -Members $target -ErrorAction Stop
            Write-Host "Toegevoegd aan: $($group.Name)" -ForegroundColor Green
            $success++
        } catch {
            Write-Host "Mislukt bij $($group.Name): $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        }
    }

    Write-Host "`nKlaar. Succesvol: $success | Mislukt: $failed" -ForegroundColor Cyan
} else {
    Write-Host "`nGeannuleerd. Er is niets gewijzigd." -ForegroundColor Yellow
}
