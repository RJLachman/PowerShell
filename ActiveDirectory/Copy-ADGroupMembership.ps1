<#
.SYNOPSIS
    Kopieert de groepslidmaatschappen van een brongebruiker naar een doelgebruiker in on-premises AD.

.DESCRIPTION
    Zoekt de bron- en doelgebruiker op aan de hand van hun e-mailadres (attribuut 'mail'),
    bepaalt welke groepen de bron heeft die het doel nog niet heeft en toont dat overzicht.
    Na een J/N-bevestiging wordt de doelgebruiker aan die groepen toegevoegd.

    Het script is idempotent: groepen waar het doel al lid van is worden overgeslagen.
    Met -WhatIf zie je welke groepen toegevoegd zouden worden, zonder iets te wijzigen.
    Met -Force wordt de J/N-vraag overgeslagen.

.PARAMETER SourceEmail
    E-mailadres van de brongebruiker. Wordt gevraagd als je het weglaat.

.PARAMETER TargetEmail
    E-mailadres van de doelgebruiker. Wordt gevraagd als je het weglaat.

.PARAMETER Force
    Slaat de J/N-bevestiging over.

.EXAMPLE
    .\Copy-ADGroupMembership.ps1

    Vraagt om bron en doel, toont de groepen en vraagt (J/N) om bevestiging.

.EXAMPLE
    .\Copy-ADGroupMembership.ps1 -SourceEmail 'jan@contoso.nl' -TargetEmail 'piet@contoso.nl' -WhatIf

    Toont welke groepen gekopieerd zouden worden, zonder iets te wijzigen.

.EXAMPLE
    .\Copy-ADGroupMembership.ps1 -SourceEmail 'jan@contoso.nl' -TargetEmail 'piet@contoso.nl' -Force

    Kopieert de groepen zonder J/N-vraag.

.NOTES
    Auteur   : RJ Lachman
    Versie   : 1.2.0
    Vereisten: Module ActiveDirectory (RSAT) en rechten om groepsleden te wijzigen
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [Parameter()]
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$SourceEmail,

    [Parameter()]
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$TargetEmail,

    [Parameter()]
    [switch]$Force
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    Write-Host "Module ActiveDirectory (RSAT) kon niet worden geladen: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# --- Invoer -------------------------------------------------------------------
if (-not $SourceEmail) { $SourceEmail = (Read-Host 'Geef het e-mailadres van de brongebruiker op').Trim() }
if (-not $TargetEmail) { $TargetEmail = (Read-Host 'Geef het e-mailadres van de doelgebruiker op').Trim() }

if ([string]::IsNullOrWhiteSpace($SourceEmail) -or [string]::IsNullOrWhiteSpace($TargetEmail)) {
    Write-Host 'Bron of doel is leeg. Script stopt.' -ForegroundColor Red
    return
}

if ($SourceEmail -eq $TargetEmail) {
    Write-Host 'Bron en doel zijn hetzelfde. Script stopt.' -ForegroundColor Red
    return
}

# --- Gebruikers ophalen -------------------------------------------------------
Write-Host "`nGebruikers ophalen..." -ForegroundColor Yellow
# Enkele quotes escapen zodat een adres het AD-filter niet kan breken.
$sourceFilter = "mail -eq '$($SourceEmail.Replace("'", "''"))'"
$targetFilter = "mail -eq '$($TargetEmail.Replace("'", "''"))'"
try {
    $source = @(Get-ADUser -Filter $sourceFilter -Properties MemberOf, mail -ErrorAction Stop)
    $target = @(Get-ADUser -Filter $targetFilter -Properties MemberOf, mail -ErrorAction Stop)
}
catch {
    Write-Host "Fout bij ophalen gebruiker: $($_.Exception.Message)" -ForegroundColor Red
    return
}

if ($source.Count -eq 0) {
    Write-Host "Brongebruiker met e-mail '$SourceEmail' niet gevonden." -ForegroundColor Red
    return
}
if ($target.Count -eq 0) {
    Write-Host "Doelgebruiker met e-mail '$TargetEmail' niet gevonden." -ForegroundColor Red
    return
}
if ($source.Count -gt 1) {
    Write-Host "Meerdere gebruikers gevonden met e-mail '$SourceEmail'. Script stopt." -ForegroundColor Red
    return
}
if ($target.Count -gt 1) {
    Write-Host "Meerdere gebruikers gevonden met e-mail '$TargetEmail'. Script stopt." -ForegroundColor Red
    return
}
$source = $source[0]
$target = $target[0]

Write-Host "Bron: $($source.Name) <$($source.mail)>" -ForegroundColor Green
Write-Host "Doel: $($target.Name) <$($target.mail)>" -ForegroundColor Green

# --- Groepen bepalen ----------------------------------------------------------
$sourceGroups = @($source.MemberOf)
if ($sourceGroups.Count -eq 0) {
    Write-Host "`nDe brongebruiker zit in geen enkele groep." -ForegroundColor Yellow
    return
}

$targetGroups = @($target.MemberOf)
$groupsToAdd = @($sourceGroups | Where-Object { $_ -notin $targetGroups })

if ($groupsToAdd.Count -eq 0) {
    Write-Host "`nDe doelgebruiker zit al in alle groepen van de bron." -ForegroundColor Yellow
    return
}

Write-Host "`nDe volgende groepen worden gekopieerd ($($groupsToAdd.Count) stuks):" -ForegroundColor Cyan
Write-Host '----------------------------------------'
$groups = foreach ($groupDN in $groupsToAdd) {
    try {
        $group = Get-ADGroup -Identity $groupDN -ErrorAction Stop
        Write-Host "- $($group.Name)" -ForegroundColor White
        $group
    }
    catch {
        Write-Host "- $groupDN (kan groep niet ophalen: $($_.Exception.Message))" -ForegroundColor Red
    }
}
Write-Host "----------------------------------------`n"

# --- Bevestiging --------------------------------------------------------------
if (-not $WhatIfPreference -and -not $Force) {
    $confirm = Read-Host 'Weet je het zeker? Deze memberships kopiëren? (J/N)'
    if ($confirm -notin @('J', 'j', 'Ja', 'ja')) {
        Write-Host "`nGeannuleerd. Er is niets gewijzigd." -ForegroundColor Yellow
        return
    }
}
# De J/N-vraag (of -Force) is de bevestiging; geen tweede prompt per groep tenzij -Confirm is opgegeven.
if (-not $PSBoundParameters.ContainsKey('Confirm')) { $ConfirmPreference = 'None' }

# --- Kopiëren -----------------------------------------------------------------
if (-not $WhatIfPreference) { Write-Host "`nKopiëren starten..." -ForegroundColor Yellow }
$success = 0
$failed = 0

foreach ($group in $groups) {
    if (-not $PSCmdlet.ShouldProcess("Groep '$($group.Name)'", "Lid toevoegen: $($target.Name)")) { continue }
    try {
        Add-ADGroupMember -Identity $group -Members $target -ErrorAction Stop
        Write-Host "Toegevoegd aan: $($group.Name)" -ForegroundColor Green
        $success++
    }
    catch {
        Write-Host "Mislukt bij $($group.Name): $($_.Exception.Message)" -ForegroundColor Red
        $failed++
    }
}

if (-not $WhatIfPreference) {
    Write-Host "`nKlaar. Succesvol: $success | Mislukt: $failed" -ForegroundColor Cyan
}
