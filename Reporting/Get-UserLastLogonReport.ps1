<#
.SYNOPSIS
    Maakt een overzichtelijk CSV-rapport van de laatste aanmelding van gebruikers.

.DESCRIPTION
    Leest e-mailadressen uit een tekstbestand (één adres per regel) en zoekt elke
    gebruiker op in Entra ID (Microsoft Graph) en/of on-premises Active Directory.
    Per gebruiker worden naam, afdeling, functie, accountstatus en de laatste
    aanmelding opgehaald. Het resultaat wordt geëxporteerd naar een CSV-bestand
    (puntkomma als scheidingsteken, zodat het direct goed opent in Nederlandse Excel).

    Bronnen voor de laatste aanmelding:
      - Entra ID : signInActivity (interactief, niet-interactief en laatste geslaagde)
                   Vereist Entra ID P1/P2 en de scopes User.Read.All + AuditLog.Read.All.
      - AD       : LastLogonTimestamp (gerepliceerd, kan 9 tot 14 dagen achterlopen).
                   Met -ExactADLogon wordt LastLogon op alle domeincontrollers opgevraagd.

.PARAMETER InputFile
    Tekstbestand met e-mailadressen. Lege regels en regels die met # beginnen worden genegeerd.

.PARAMETER OutputFolder
    Map waarin het CSV-rapport wordt opgeslagen.

.PARAMETER Source
    ActiveDirectory (standaard), EntraID of Both.

.PARAMETER InactiveDays
    Na hoeveel dagen zonder aanmelding een gebruiker als 'Inactief' wordt gemarkeerd (standaard 30).

.PARAMETER ExactADLogon
    Vraagt LastLogon op bij elke domeincontroller voor de exacte AD-aanmelding (trager).

.EXAMPLE
    .\Get-UserLastLogonReport.ps1

.EXAMPLE
    .\Get-UserLastLogonReport.ps1 -Source EntraID -InactiveDays 60

.EXAMPLE
    .\Get-UserLastLogonReport.ps1 -InputFile 'D:\lijsten\users.txt' -OutputFolder 'D:\rapporten' -ExactADLogon
#>
[CmdletBinding()]
param(
    [string]$InputFile = 'C:\temp\userslogin.txt',
    [string]$OutputFolder = 'C:\temp',
    [ValidateSet('Both', 'EntraID', 'ActiveDirectory')]
    [string]$Source = 'ActiveDirectory',
    [ValidateRange(1, 3650)]
    [int]$InactiveDays = 30,
    [switch]$ExactADLogon
)

$ErrorActionPreference = 'Stop'
$useEntra = $Source -in 'Both', 'EntraID'
$useAD    = $Source -in 'Both', 'ActiveDirectory'
$dateFmt  = 'dd-MM-yyyy HH:mm'

function Write-Status([string]$Message, [string]$Color = 'Cyan') {
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $Message" -ForegroundColor $Color
}

function Format-Date($Date) {
    if ($Date) { return $Date.ToString($dateFmt) } else { return '' }
}

# --- Invoer inlezen -----------------------------------------------------------
if (-not (Test-Path -LiteralPath $InputFile)) {
    Write-Status "Invoerbestand niet gevonden: $InputFile" 'Red'
    return
}

$emailRegex = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
$rawLines   = Get-Content -LiteralPath $InputFile |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ -and -not $_.StartsWith('#') }

$emails  = $rawLines | Where-Object { $_ -match $emailRegex } | Sort-Object -Unique
$invalid = $rawLines | Where-Object { $_ -notmatch $emailRegex }

if (-not $emails) {
    Write-Status "Geen geldige e-mailadressen gevonden in $InputFile" 'Red'
    return
}
Write-Status "$(@($emails).Count) unieke e-mailadressen ingelezen uit $InputFile"
if ($invalid) {
    Write-Status "$(@($invalid).Count) ongeldige regel(s) overgeslagen: $($invalid -join ', ')" 'Yellow'
}

# --- Verbinden ----------------------------------------------------------------
if ($useEntra) {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
        Write-Status 'Module Microsoft.Graph.Users ontbreekt (Install-Module Microsoft.Graph -Scope CurrentUser). Entra ID wordt overgeslagen.' 'Yellow'
        $useEntra = $false
    }
}
if ($useEntra) {
    Import-Module Microsoft.Graph.Users
    $scopes = 'User.Read.All', 'AuditLog.Read.All'
    try {
        # WAM (Windows-aanmeldvenster) uitschakelen; voorkomt 'window handle'-fouten
        if (Get-Command Set-MgGraphOption -ErrorAction SilentlyContinue) {
            Set-MgGraphOption -DisableLoginByWAM $true
        }
        Connect-MgGraph -Scopes $scopes -NoWelcome
    }
    catch {
        Write-Status 'Interactieve login mislukt, terugvallen op device code...' 'Yellow'
        Connect-MgGraph -Scopes $scopes -UseDeviceCode -NoWelcome
    }
    Write-Status "Verbonden met Microsoft Graph als $((Get-MgContext).Account)" 'Green'
}

if ($useAD) {
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-Status 'Module ActiveDirectory ontbreekt (RSAT). AD wordt overgeslagen.' 'Yellow'
        $useAD = $false
    }
    else {
        Import-Module ActiveDirectory
        if ($ExactADLogon) {
            $domainControllers = (Get-ADDomainController -Filter *).HostName
            Write-Status "Exacte AD-logon: $(@($domainControllers).Count) domeincontroller(s) worden bevraagd"
        }
    }
}

if (-not $useEntra -and -not $useAD) {
    Write-Status 'Geen bron beschikbaar om te bevragen. Gestopt.' 'Red'
    return
}

# --- Gebruikers ophalen -------------------------------------------------------
$graphProps = 'id', 'displayName', 'givenName', 'surname', 'mail', 'userPrincipalName',
              'department', 'jobTitle', 'accountEnabled', 'createdDateTime', 'signInActivity'
$adProps    = 'DisplayName', 'GivenName', 'Surname', 'mail', 'UserPrincipalName', 'Department',
              'Title', 'Enabled', 'LastLogonTimestamp', 'whenCreated', 'SamAccountName'

$now     = Get-Date
$results = [System.Collections.Generic.List[object]]::new()
$i       = 0

foreach ($email in $emails) {
    $i++
    Write-Progress -Activity 'Gebruikers opzoeken' -Status "$email ($i van $(@($emails).Count))" `
                   -PercentComplete (($i / @($emails).Count) * 100)

    $safe       = $email.Replace("'", "''")
    $entraUser  = $null
    $adUser     = $null
    $entraLogon = $null
    $adLogon    = $null
    $notes      = [System.Collections.Generic.List[string]]::new()

    if ($useEntra) {
        try {
            $entraUser = Get-MgUser -Filter "userPrincipalName eq '$safe' or mail eq '$safe'" `
                                    -Property $graphProps -ErrorAction Stop | Select-Object -First 1
            if ($entraUser.SignInActivity) {
                $sia = $entraUser.SignInActivity
                $entraLogon = @($sia.LastSignInDateTime, $sia.LastNonInteractiveSignInDateTime,
                                $sia.LastSuccessfulSignInDateTime) |
                              Where-Object { $_ } | ForEach-Object { ([datetime]$_).ToLocalTime() } |
                              Sort-Object -Descending | Select-Object -First 1
            }
        }
        catch { $notes.Add("Entra-fout: $($_.Exception.Message)") }
    }

    if ($useAD) {
        try {
            $adUser = Get-ADUser -Filter "mail -eq '$safe' -or UserPrincipalName -eq '$safe'" `
                                 -Properties $adProps | Select-Object -First 1
            if ($adUser) {
                if ($ExactADLogon) {
                    $adLogon = foreach ($dc in $domainControllers) {
                        try {
                            $v = (Get-ADUser -Identity $adUser.SamAccountName -Server $dc -Properties LastLogon).LastLogon
                            if ($v -gt 0) { [datetime]::FromFileTime($v) }
                        }
                        catch { $notes.Add("DC $dc onbereikbaar") }
                    }
                    $adLogon = $adLogon | Sort-Object -Descending | Select-Object -First 1
                }
                elseif ($adUser.LastLogonTimestamp -gt 0) {
                    $adLogon = [datetime]::FromFileTime($adUser.LastLogonTimestamp)
                }
            }
        }
        catch { $notes.Add("AD-fout: $($_.Exception.Message)") }
    }

    $found     = [bool]($entraUser -or $adUser)
    $lastLogon = @($entraLogon, $adLogon) | Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1
    $daysAgo   = if ($lastLogon) { [int][math]::Floor(($now - $lastLogon).TotalDays) } else { $null }

    $enabled = if ($entraUser) { $entraUser.AccountEnabled } elseif ($adUser) { $adUser.Enabled } else { $null }

    $status = if (-not $found)                { 'Niet gevonden' }
              elseif ($enabled -eq $false)    { 'Uitgeschakeld' }
              elseif (-not $lastLogon)        { 'Nooit aangemeld' }
              elseif ($daysAgo -gt $InactiveDays) { "Inactief (> $InactiveDays dagen)" }
              else                            { 'Actief' }

    $pick = { param($a, $b) if ($a) { $a } else { $b } }

    $results.Add([pscustomobject][ordered]@{
        'Weergavenaam'              = & $pick $entraUser.DisplayName $adUser.DisplayName
        'Voornaam'                  = & $pick $entraUser.GivenName   $adUser.GivenName
        'Achternaam'                = & $pick $entraUser.Surname     $adUser.Surname
        'E-mailadres'               = $email
        'UPN'                       = & $pick $entraUser.UserPrincipalName $adUser.UserPrincipalName
        'Afdeling'                  = & $pick $entraUser.Department  $adUser.Department
        'Functie'                   = & $pick $entraUser.JobTitle    $adUser.Title
        'Account actief'            = if ($null -eq $enabled) { '' } elseif ($enabled) { 'Ja' } else { 'Nee' }
        'Status'                    = $status
        'Laatste aanmelding'        = Format-Date $lastLogon
        'Dagen geleden'             = $daysAgo
        'Laatste aanmelding Entra'  = Format-Date $entraLogon
        'Laatste aanmelding AD'     = Format-Date $adLogon
        'Aangemaakt op'             = Format-Date (& $pick $entraUser.CreatedDateTime $adUser.whenCreated)
        'Gevonden in'               = (@(if ($entraUser) { 'Entra ID' }; if ($adUser) { 'AD' }) -join ' + ')
        'Opmerkingen'               = $notes -join '; '
    })
}
Write-Progress -Activity 'Gebruikers opzoeken' -Completed

# --- Exporteren ---------------------------------------------------------------
if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}
$outFile = Join-Path $OutputFolder ("UserLastLogonReport_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmm'))

# Meest inactieve gebruikers bovenaan, niet-gevonden accounts onderaan
$sorted = $results | Sort-Object @{ Expression = { $_.Status -eq 'Niet gevonden' } },
                                 @{ Expression = { if ($null -eq $_.'Dagen geleden') { [int]::MaxValue } else { $_.'Dagen geleden' } }; Descending = $true },
                                 'Weergavenaam'

$encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
$sorted | Export-Csv -LiteralPath $outFile -NoTypeInformation -Delimiter ';' -Encoding $encoding

# --- Samenvatting -------------------------------------------------------------
Write-Host ''
Write-Host '================ SAMENVATTING ================' -ForegroundColor White
$results | Group-Object Status | Sort-Object Count -Descending | ForEach-Object {
    $color = switch -Wildcard ($_.Name) {
        'Actief'        { 'Green' }
        'Inactief*'     { 'Yellow' }
        'Niet gevonden' { 'Red' }
        default         { 'Gray' }
    }
    Write-Host ("  {0,-28} {1,5}" -f $_.Name, $_.Count) -ForegroundColor $color
}
Write-Host ("  {0,-28} {1,5}" -f 'Totaal', $results.Count) -ForegroundColor White
Write-Host '==============================================' -ForegroundColor White
Write-Status "Rapport opgeslagen: $outFile" 'Green'

$answer = Read-Host 'Rapport nu openen? (J/N)'
if ($answer -in 'J', 'j', 'Ja', 'ja') { Invoke-Item -LiteralPath $outFile }
