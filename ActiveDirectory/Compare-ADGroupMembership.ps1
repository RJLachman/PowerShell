<#
.SYNOPSIS
    Vergelijkt de leden van een on-premises AD-groep met een lijst e-mailadressen en maakt een Excel-rapport.

.DESCRIPTION
    Leest e-mailadressen uit een tekstbestand (één adres per regel) en haalt de
    gebruikersleden van een Active Directory-groep op. Het script bepaalt daarna:

      - welke adressen van de lijst nog geen lid zijn van de groep (toevoegen);
      - welke adressen van de lijst nergens in AD terug te vinden zijn;
      - welke groepsleden niet op de lijst staan;
      - welke adressen al overeenkomen met een groepslid.

    Een adres uit de lijst geldt als overeenkomst als het (hoofdletterongevoelig)
    overeenkomt met het mail-attribuut, de UserPrincipalName of een smtp:-proxyAddress
    van een groepslid.

    Het rapport wordt als Excel-bestand (.xlsx) weggeschreven als de module
    ImportExcel beschikbaar is (of na toestemming wordt geïnstalleerd). Zonder
    ImportExcel vallen de resultaten terug op aparte CSV-bestanden (puntkomma,
    UTF-8 met BOM).

    Dit script is alleen-lezen: het wijzigt de groep of gebruikersaccounts niet.
    Optioneel wordt een tekstbestand met kant-en-klare Add-ADGroupMember-commando's
    weggeschreven; die commando's worden nooit uitgevoerd.

.PARAMETER GroupName
    Naam (SamAccountName of DistinguishedName) van de AD-groep die vergeleken wordt.

.PARAMETER InputFile
    Tekstbestand met e-mailadressen. Lege regels en regels die met # beginnen worden genegeerd.
    Dubbele adressen worden één keer meegenomen en in het rapport vermeld.

.PARAMETER OutputFolder
    Map waarin het rapport (en eventueel de commando-lijst) wordt opgeslagen.

.PARAMETER Recursive
    Neemt ook leden van geneste groepen mee (Get-ADGroupMember -Recursive).

.PARAMETER Server
    Optionele domeincontroller waarop de AD-query's worden uitgevoerd.

.PARAMETER ExpectedCount
    Optioneel verwacht aantal gebruikersleden in de groep. Wijkt het werkelijke
    aantal hiervan af, dan verschijnt een waarschuwing (het script stopt niet).

.PARAMETER SkipCommandFile
    Schrijft geen tekstbestand met Add-ADGroupMember-commando's weg.

.EXAMPLE
    .\Compare-ADGroupMembership.ps1

    Gebruikt de standaardgroep soft-apps-154 en de lijst C:\temp\suwi.txt.
    Het rapport komt in C:\temp.

.EXAMPLE
    .\Compare-ADGroupMembership.ps1 -GroupName 'soft-apps-154' -InputFile 'C:\temp\suwi.txt' -ExpectedCount 84

    Vergelijkt de groep met de lijst en waarschuwt als het aantal leden niet 84 is.

.EXAMPLE
    .\Compare-ADGroupMembership.ps1 -GroupName 'APP-Finance' -InputFile 'D:\lijsten\finance.txt' -OutputFolder 'D:\rapporten' -Recursive -Server 'DC01.contoso.local'

    Inclusief geneste groepsleden, bevraagd op een specifieke domeincontroller.

.INPUTS
    Geen. Het script leest de adressen uit -InputFile.

.OUTPUTS
    Een Excel-bestand ADGroepVergelijking_<groep>_<yyyyMMdd_HHmm>.xlsx in -OutputFolder,
    of (zonder ImportExcel) meerdere CSV-bestanden plus een optioneel tekstbestand
    met Add-ADGroupMember-commando's.

.NOTES
    Auteur   : RJ Lachman
    Versie   : 1.0.0
    Vereisten: Module ActiveDirectory (RSAT). Voor een Excel-rapport: module ImportExcel
               (Install-Module ImportExcel -Scope CurrentUser).
    Dit script wijzigt niets in Active Directory; het leest alleen gegevens en schrijft een rapport.
#>
[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$GroupName = 'soft-apps-154',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$InputFile = 'C:\temp\suwi.txt',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputFolder = 'C:\temp',

    [Parameter()]
    [switch]$Recursive,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$Server,

    [Parameter()]
    [ValidateRange(0, 1000000)]
    [int]$ExpectedCount,

    [Parameter()]
    [switch]$SkipCommandFile
)

$ErrorActionPreference = 'Stop'
$script:GeneratedAt = Get-Date
$script:ExpectedCountOpgegeven = $PSBoundParameters.ContainsKey('ExpectedCount')

function Write-Status {
    <#
    .SYNOPSIS
        Schrijft een statusregel met tijdstempel naar de console.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter()]
        [string]$Color = 'Cyan'
    )
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $Message" -ForegroundColor $Color
}

function Get-SafeAdFilterValue {
    <#
    .SYNOPSIS
        Escapet enkele aanhalingstekens voor gebruik in een AD-filter.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )
    return $Value.Replace("'", "''")
}

function Get-NormalizedEmail {
    <#
    .SYNOPSIS
        Normaliseert een e-mailadres (trim + lowercase) voor vergelijking.
    #>
    param(
        [Parameter()]
        [string]$Value
    )
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    return $Value.Trim().ToLowerInvariant()
}

function Get-ProxySmtpAddress {
    <#
    .SYNOPSIS
        Haalt de smtp-adressen uit proxyAddresses en normaliseert ze.
    #>
    param(
        [Parameter()]
        $ProxyAddresses
    )
    $result = [System.Collections.Generic.List[string]]::new()
    if (-not $ProxyAddresses) { return @() }
    foreach ($entry in @($ProxyAddresses)) {
        if ([string]::IsNullOrWhiteSpace($entry)) { continue }
        $text = [string]$entry
        if ($text.StartsWith('smtp:', [System.StringComparison]::OrdinalIgnoreCase)) {
            $normalized = Get-NormalizedEmail ($text.Substring(5))
            if ($normalized) { $result.Add($normalized) }
        }
    }
    return @($result | Sort-Object -Unique)
}

function Test-EmailFormat {
    <#
    .SYNOPSIS
        Controleert of een waarde op een eenvoudig e-mailadres lijkt.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )
    return $Value -match '^[^@\s]+@[^@\s]+\.[^@\s]+$'
}

function Get-AdServerParameter {
    <#
    .SYNOPSIS
        Geeft een hashtable met -Server terug als die parameter is opgegeven.
    #>
    $common = @{}
    if ($Server) { $common['Server'] = $Server }
    return $common
}

function ConvertTo-UserRecord {
    <#
    .SYNOPSIS
        Zet een AD-gebruiker om naar een intern record met genormaliseerde adressen.
    #>
    param(
        [Parameter(Mandatory)]
        $AdUser,

        [Parameter()]
        [string]$ListEmail
    )

    $mail = Get-NormalizedEmail $AdUser.mail
    $upn  = Get-NormalizedEmail $AdUser.UserPrincipalName
    $proxies = Get-ProxySmtpAddress $AdUser.proxyAddresses
    $all = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($addr in @($mail, $upn) + @($proxies)) {
        if ($addr) { [void]$all.Add($addr) }
    }

    [pscustomobject]@{
        DistinguishedName = $AdUser.DistinguishedName
        ObjectGuid        = [string]$AdUser.ObjectGuid
        Weergavenaam      = $AdUser.DisplayName
        Voornaam          = $AdUser.GivenName
        Achternaam        = $AdUser.Surname
        SamAccountName    = $AdUser.SamAccountName
        Mail              = $AdUser.mail
        UserPrincipalName = $AdUser.UserPrincipalName
        ProxyAdressen     = ($proxies -join '; ')
        Afdeling          = $AdUser.Department
        Functie           = $AdUser.Title
        AccountActief     = if ($null -eq $AdUser.Enabled) { $null } else { [bool]$AdUser.Enabled }
        Adressen          = $all
        LijstAdres        = $ListEmail
    }
}

function Find-AdUserByEmail {
    <#
    .SYNOPSIS
        Zoekt een AD-gebruiker op mail, UPN of smtp-proxyAddress.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Email,

        [Parameter()]
        [hashtable]$AdParams
    )

    $safe  = Get-SafeAdFilterValue $Email
    $props = $script:UserProperties
    $filter = "mail -eq '$safe' -or UserPrincipalName -eq '$safe'"

    try {
        $hits = @(Get-ADUser -Filter $filter -Properties $props @AdParams -ErrorAction Stop)
    }
    catch {
        Write-Verbose "Zoeken op mail/UPN voor $Email mislukt: $($_.Exception.Message)"
        $hits = @()
    }

    if ($hits.Count -eq 0) {
        # LDAP-filter: escaping van \ * ( ) en NUL is voldoende voor een e-mailadres.
        $ldapSafe = $Email.Replace('\', '\5c').Replace('*', '\2a').Replace('(', '\28').Replace(')', '\29').Replace("`0", '\00')
        $ldapFilter = "(&(objectCategory=person)(objectClass=user)(|(mail=$ldapSafe)(userPrincipalName=$ldapSafe)(proxyAddresses=smtp:$ldapSafe)(proxyAddresses=SMTP:$ldapSafe)))"
        try {
            $hits = @(Get-ADUser -LDAPFilter $ldapFilter -Properties $props @AdParams -ErrorAction Stop)
        }
        catch {
            Write-Verbose "Zoeken via LDAP/proxyAddresses voor $Email mislukt: $($_.Exception.Message)"
            $hits = @()
        }
    }

    return @($hits)
}

function Get-CsvEncodingName {
    <#
    .SYNOPSIS
        Kiest de juiste Encoding-naam voor UTF-8 met BOM, afhankelijk van de PowerShell-versie.
    #>
    if ($PSVersionTable.PSVersion.Major -ge 6) { return 'utf8BOM' }
    return 'UTF8'
}

function ConvertTo-JaNee {
    <#
    .SYNOPSIS
        Zet een boolean om naar Ja/Nee voor het rapport.
    #>
    param($Value)
    if ($null -eq $Value) { return '' }
    if ([bool]$Value) { return 'Ja' }
    return 'Nee'
}

function Get-ReportRow {
    <#
    .SYNOPSIS
        Bouwt een rapportregel met vaste kolommen in het Nederlands.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Categorie,

        [Parameter()]
        [string]$LijstAdres = '',

        [Parameter()]
        $User
    )

    $actief = ''
    if ($User) { $actief = ConvertTo-JaNee $User.AccountActief }

    [pscustomobject]@{
        Categorie         = $Categorie
        'E-mail op lijst' = $LijstAdres
        Weergavenaam      = if ($User) { $User.Weergavenaam } else { '' }
        Voornaam          = if ($User) { $User.Voornaam } else { '' }
        Achternaam        = if ($User) { $User.Achternaam } else { '' }
        SamAccountName    = if ($User) { $User.SamAccountName } else { '' }
        'E-mail (mail)'   = if ($User) { $User.Mail } else { '' }
        UPN               = if ($User) { $User.UserPrincipalName } else { '' }
        ProxyAdressen     = if ($User) { $User.ProxyAdressen } else { '' }
        Afdeling          = if ($User) { $User.Afdeling } else { '' }
        Functie           = if ($User) { $User.Functie } else { '' }
        'Account actief'  = $actief
    }
}

function Test-ImportExcelAvailable {
    <#
    .SYNOPSIS
        Geeft terug of de module ImportExcel geladen kan worden.
    #>
    return [bool](Get-Module -ListAvailable -Name ImportExcel)
}

function Initialize-ImportExcelModule {
    <#
    .SYNOPSIS
        Biedt aan om ImportExcel te installeren en geeft daarna de beschikbaarheid terug.
    #>
    Write-Status 'Module ImportExcel ontbreekt. Zonder deze module wordt het rapport als CSV weggeschreven.' 'Yellow'
    Write-Host "  Installeer met: Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Yellow
    $answer = Read-Host 'ImportExcel nu installeren? (J/N)'
    if ($answer -notin 'J', 'j', 'Ja', 'ja') {
        Write-Status 'Installatie overgeslagen. Er worden CSV-bestanden gemaakt.' 'Yellow'
        return $false
    }

    try {
        Write-Status 'ImportExcel wordt geïnstalleerd...'
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Import-Module ImportExcel -ErrorAction Stop
        Write-Status 'ImportExcel is geïnstalleerd.' 'Green'
        return $true
    }
    catch {
        Write-Status "Installatie van ImportExcel is mislukt: $($_.Exception.Message). Er worden CSV-bestanden gemaakt." 'Yellow'
        return $false
    }
}

function Export-CsvSafe {
    <#
    .SYNOPSIS
        Exporteert objecten naar een CSV met puntkomma en UTF-8 BOM.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter()]
        $Data
    )
    $rows = @($Data)
    if ($rows.Count -eq 0) { return }
    $rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Delimiter ';' -Encoding (Get-CsvEncodingName)
}

function Add-WorksheetIfData {
    <#
    .SYNOPSIS
        Voegt een werkblad toe, inclusief tabelopmaak. Lege categorieën krijgen een meldingrij.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$ExcelPath,

        [Parameter(Mandatory)]
        [string]$WorksheetName,

        [Parameter()]
        $Data,

        [Parameter()]
        [string]$TableStyle = 'Medium2'
    )
    $rows = @($Data)
    if ($rows.Count -eq 0) {
        $rows = @(
            [pscustomobject]@{ Opmerking = 'Geen rijen in deze categorie.' }
        )
    }

    $params = @{
        Path          = $ExcelPath
        WorksheetName = $WorksheetName
        AutoSize      = $true
        AutoFilter    = $true
        FreezeTopRow  = $true
        BoldTopRow    = $true
        TableName     = ($WorksheetName -replace '[^A-Za-z0-9]', '')
        TableStyle    = $TableStyle
        PassThru      = $true
    }
    if (Test-Path -LiteralPath $ExcelPath) {
        $params['Append'] = $true
    }

    $excel = $rows | Export-Excel @params
    try {
        $ws = $excel.Workbook.Worksheets[$WorksheetName]
        if ($ws -and $ws.Dimension -and $Data -and @($Data).Count -gt 0) {
            try {
                $endRow = $ws.Dimension.End.Row
                $fill = $null
                switch ($WorksheetName) {
                    'Toevoegen aan groep'  { $fill = [System.Drawing.Color]::FromArgb(255, 230, 153) }
                    'Niet op lijst'        { $fill = [System.Drawing.Color]::FromArgb(189, 215, 238) }
                    'Overeenkomend'        { $fill = [System.Drawing.Color]::FromArgb(198, 239, 206) }
                    'Niet gevonden in AD'  { $fill = [System.Drawing.Color]::FromArgb(255, 199, 206) }
                }
                if ($fill) {
                    $range = $ws.Cells["A2:A$endRow"]
                    $range.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $range.Style.Fill.BackgroundColor.SetColor($fill)
                }
            }
            catch {
                Write-Verbose "Voorwaardelijke opmaak voor werkblad $WorksheetName overgeslagen: $($_.Exception.Message)"
            }
        }
        $excel.Save()
    }
    finally {
        $excel.Dispose()
    }
}

# --- Modules ------------------------------------------------------------------
try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    Write-Status "Module ActiveDirectory (RSAT) kon niet worden geladen: $($_.Exception.Message)" 'Red'
    return
}

$script:UserProperties = @(
    'DisplayName', 'GivenName', 'Surname', 'mail', 'UserPrincipalName',
    'SamAccountName', 'proxyAddresses', 'Department', 'Title', 'Enabled',
    'DistinguishedName', 'ObjectGuid'
)
$adParams = Get-AdServerParameter

# --- Invoer inlezen -----------------------------------------------------------
if (-not (Test-Path -LiteralPath $InputFile)) {
    Write-Status "Invoerbestand niet gevonden: $InputFile" 'Red'
    return
}

Write-Verbose "Invoerbestand lezen: $InputFile"
$rawLines = Get-Content -LiteralPath $InputFile
$parsedLines = [System.Collections.Generic.List[object]]::new()
$lineNumber = 0
foreach ($raw in @($rawLines)) {
    $lineNumber++
    $trimmed = if ($null -eq $raw) { '' } else { $raw.Trim() }
    if (-not $trimmed -or $trimmed.StartsWith('#')) { continue }
    $parsedLines.Add([pscustomobject]@{
        Regelnummer = $lineNumber
        Oorspronkelijk = $trimmed
        Genormaliseerd = (Get-NormalizedEmail $trimmed)
        Geldig = (Test-EmailFormat $trimmed)
    })
}

$invalidEntries = @($parsedLines | Where-Object { -not $_.Geldig })
$validEntries   = @($parsedLines | Where-Object { $_.Geldig })

$seen = @{}
$emails = [System.Collections.Generic.List[string]]::new()
$duplicateEntries = [System.Collections.Generic.List[object]]::new()
foreach ($entry in $validEntries) {
    $key = $entry.Genormaliseerd
    if ($seen.ContainsKey($key)) {
        $duplicateEntries.Add([pscustomobject]@{
            Regelnummer      = $entry.Regelnummer
            Waarde           = $entry.Oorspronkelijk
            Reden            = "Duplicaat van regel $($seen[$key].Regelnummer)"
            Genormaliseerd   = $key
        })
        continue
    }
    $seen[$key] = $entry
    $emails.Add($key)
}

if ($emails.Count -eq 0) {
    Write-Status "Geen geldige e-mailadressen gevonden in $InputFile" 'Red'
    return
}

Write-Status "$($emails.Count) unieke e-mailadressen ingelezen uit $InputFile"
if ($invalidEntries) {
    Write-Status "$($invalidEntries.Count) ongeldige regel(s) overgeslagen." 'Yellow'
}
if ($duplicateEntries) {
    Write-Status "$($duplicateEntries.Count) dubbele regel(s) overgeslagen." 'Yellow'
}

# --- Groep ophalen ------------------------------------------------------------
Write-Status "AD-groep ophalen: $GroupName"
try {
    $group = Get-ADGroup -Identity $GroupName -Properties DistinguishedName, SamAccountName, Name, GroupScope, GroupCategory, Description @adParams -ErrorAction Stop
}
catch {
    Write-Status "AD-groep '$GroupName' niet gevonden of niet leesbaar: $($_.Exception.Message)" 'Red'
    return
}

Write-Verbose "Groep gevonden: $($group.DistinguishedName)"

$memberIds = @()
$usedMemberFallback = $false
$memberParams = @{} + $adParams
if ($Recursive) { $memberParams['Recursive'] = $true }

try {
    Write-Verbose 'Groepsleden ophalen met Get-ADGroupMember.'
    $memberIds = @(Get-ADGroupMember -Identity $group.DistinguishedName @memberParams -ErrorAction Stop)
}
catch {
    Write-Status "Get-ADGroupMember mislukte (mogelijk de limiet van 5000 leden): $($_.Exception.Message). Terugval via het Member-attribuut." 'Yellow'
    $usedMemberFallback = $true
    try {
        $groupWithMembers = Get-ADGroup -Identity $group.DistinguishedName -Properties Member @adParams -ErrorAction Stop
        $memberIds = @($groupWithMembers.Member)
        if ($Recursive -and $memberIds) {
            Write-Verbose 'Terugval: geneste groepen handmatig uitvouwen.'
            $queue = [System.Collections.Generic.Queue[string]]::new()
            $visited = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $expanded = [System.Collections.Generic.List[string]]::new()
            foreach ($dn in @($groupWithMembers.Member)) { $queue.Enqueue([string]$dn) }
            while ($queue.Count -gt 0) {
                $current = $queue.Dequeue()
                if (-not $visited.Add($current)) { continue }
                try {
                    $obj = Get-ADObject -Identity $current -Properties objectClass @adParams -ErrorAction Stop
                }
                catch {
                    Write-Verbose "Object $current kon niet worden gelezen: $($_.Exception.Message)"
                    continue
                }
                if ($obj.objectClass -contains 'group' -or $obj.objectClass -eq 'group') {
                    try {
                        $nested = Get-ADGroup -Identity $current -Properties Member @adParams -ErrorAction Stop
                        foreach ($nestedDn in @($nested.Member)) { $queue.Enqueue([string]$nestedDn) }
                    }
                    catch {
                        Write-Verbose "Geneste groep $current kon niet worden gelezen: $($_.Exception.Message)"
                    }
                }
                else {
                    $expanded.Add($current)
                }
            }
            $memberIds = @($expanded)
        }
    }
    catch {
        Write-Status "Groepsleden ophalen via het Member-attribuut is mislukt: $($_.Exception.Message)" 'Red'
        return
    }
}

Write-Verbose "Ruwe ledenobjecten: $(@($memberIds).Count) (terugval: $usedMemberFallback)"

# Alleen gebruikers resolven (geen computers, contactpersonen of groepen).
$groupUsers = [System.Collections.Generic.List[object]]::new()
$skippedNonUsers = 0
$failedMembers = 0
$memberIndex = 0
$rawMembers = @($memberIds)

foreach ($member in $rawMembers) {
    $memberIndex++
    if ($rawMembers.Count -gt 0) {
        Write-Progress -Activity 'Groepsleden ophalen' -Status "$memberIndex van $($rawMembers.Count)" `
                       -PercentComplete (($memberIndex / $rawMembers.Count) * 100)
    }

    $identity = $null
    $objectClass = $null
    if ($member -is [string]) {
        $identity = $member
    }
    else {
        $identity = $member.DistinguishedName
        if (-not $identity) { $identity = $member.SamAccountName }
        if ($member.PSObject.Properties['objectClass']) { $objectClass = $member.objectClass }
        elseif ($member.PSObject.Properties['ObjectClass']) { $objectClass = $member.ObjectClass }
    }

    if ($objectClass) {
        $classes = @($objectClass)
        $isComputer = $classes -contains 'computer'
        $isUser = ($classes -contains 'user' -or $classes -contains 'inetOrgPerson') -and -not $isComputer
        if (-not $isUser) {
            $skippedNonUsers++
            Write-Verbose "Lid overgeslagen (geen gebruiker): $identity"
            continue
        }
    }

    if (-not $identity) {
        $failedMembers++
        continue
    }

    try {
        $adUser = Get-ADUser -Identity $identity -Properties $script:UserProperties @adParams -ErrorAction Stop
        $groupUsers.Add((ConvertTo-UserRecord -AdUser $adUser))
    }
    catch {
        # Geen gebruiker (groep, computer, verwijderd object) of geen leesrechten.
        if ($objectClass) {
            $failedMembers++
            Write-Verbose "Gebruiker $identity kon niet worden gelezen: $($_.Exception.Message)"
        }
        else {
            try {
                $obj = Get-ADObject -Identity $identity -Properties objectClass @adParams -ErrorAction Stop
                $classes = @($obj.objectClass)
                $isComputer = $classes -contains 'computer'
                $isUser = ($classes -contains 'user' -or $classes -contains 'inetOrgPerson') -and -not $isComputer
                if ($isUser) {
                    $failedMembers++
                    Write-Verbose "Gebruiker $identity kon niet worden gelezen: $($_.Exception.Message)"
                }
                else {
                    $skippedNonUsers++
                    Write-Verbose "Lid overgeslagen (geen gebruiker): $identity ($($classes -join ','))"
                }
            }
            catch {
                $failedMembers++
                Write-Verbose "Lid $identity kon niet worden opgehaald: $($_.Exception.Message)"
            }
        }
    }
}
Write-Progress -Activity 'Groepsleden ophalen' -Completed

# Uniek op DistinguishedName (geneste groepen kunnen dezelfde gebruiker meerdere keren opleveren).
$uniqueUsers = [System.Collections.Generic.List[object]]::new()
$seenDn = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($user in $groupUsers) {
    if ($seenDn.Add($user.DistinguishedName)) {
        $uniqueUsers.Add($user)
    }
}
$groupUsers = $uniqueUsers

Write-Status "$($groupUsers.Count) gebruikerslid(leden) gevonden in groep '$($group.SamAccountName)'."
if ($skippedNonUsers -gt 0) {
    Write-Status "$skippedNonUsers lid/leden overgeslagen (geen gebruikersobject)." 'Yellow'
}
if ($failedMembers -gt 0) {
    Write-Status "$failedMembers lid/leden konden niet worden uitgelezen." 'Yellow'
}
if ($script:ExpectedCountOpgegeven -and $groupUsers.Count -ne $ExpectedCount) {
    Write-Status "Verwacht aantal leden is $ExpectedCount, maar de groep bevat $($groupUsers.Count) gebruiker(s)." 'Yellow'
}

# Index van alle adressen van groepsleden
$memberByEmail = @{}
foreach ($user in $groupUsers) {
    foreach ($addr in $user.Adressen) {
        if (-not $memberByEmail.ContainsKey($addr)) {
            $memberByEmail[$addr] = $user
        }
    }
}

# --- Vergelijken --------------------------------------------------------------
$matchedRows      = [System.Collections.Generic.List[object]]::new()
$toAddRows        = [System.Collections.Generic.List[object]]::new()
$notFoundRows     = [System.Collections.Generic.List[object]]::new()
$toAddUsers       = [System.Collections.Generic.List[object]]::new()
$matchedMemberDns = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

$listIndex = 0
foreach ($email in $emails) {
    $listIndex++
    Write-Progress -Activity 'Lijst vergelijken met de AD-groep' -Status "$email ($listIndex van $($emails.Count))" `
                   -PercentComplete (($listIndex / $emails.Count) * 100)

    if ($memberByEmail.ContainsKey($email)) {
        $user = $memberByEmail[$email]
        [void]$matchedMemberDns.Add($user.DistinguishedName)
        $matchedRows.Add((Get-ReportRow -Categorie 'Overeenkomend' -LijstAdres $email -User $user))
        continue
    }

    $hits = @(Find-AdUserByEmail -Email $email -AdParams $adParams)
    if ($hits.Count -eq 0) {
        $notFoundRows.Add((Get-ReportRow -Categorie 'Niet gevonden in AD' -LijstAdres $email))
        continue
    }

    if ($hits.Count -gt 1) {
        Write-Verbose "Meerdere AD-accounts gevonden voor $email; de eerste wordt gebruikt."
    }

    $foundUser = ConvertTo-UserRecord -AdUser $hits[0] -ListEmail $email
    $toAddUsers.Add($foundUser)
    $toAddRows.Add((Get-ReportRow -Categorie 'Moet toegevoegd worden' -LijstAdres $email -User $foundUser))
}
Write-Progress -Activity 'Lijst vergelijken met de AD-groep' -Completed

$extraRows = [System.Collections.Generic.List[object]]::new()
foreach ($user in $groupUsers) {
    if (-not $matchedMemberDns.Contains($user.DistinguishedName)) {
        $extraRows.Add((Get-ReportRow -Categorie 'Staat in groep, niet op lijst' -User $user))
    }
}

$invalidRows = [System.Collections.Generic.List[object]]::new()
foreach ($entry in $invalidEntries) {
    $invalidRows.Add([pscustomobject]@{
        Regelnummer = $entry.Regelnummer
        Waarde      = $entry.Oorspronkelijk
        Reden       = 'Ongeldig e-mailadres'
    })
}
foreach ($entry in $duplicateEntries) {
    $invalidRows.Add([pscustomobject]@{
        Regelnummer = $entry.Regelnummer
        Waarde      = $entry.Waarde
        Reden       = $entry.Reden
    })
}

# --- Uitvoermap ---------------------------------------------------------------
try {
    if (-not (Test-Path -LiteralPath $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    }
}
catch {
    Write-Status "Uitvoermap $OutputFolder kan niet worden aangemaakt: $($_.Exception.Message)" 'Red'
    return
}

$stamp = $script:GeneratedAt.ToString('yyyyMMdd_HHmm')
$safeGroupName = ($group.SamAccountName -replace '[\\/:*?"<>|]', '_')
$baseName = "ADGroepVergelijking_${safeGroupName}_$stamp"
$excelPath = Join-Path $OutputFolder "$baseName.xlsx"
$commandPath = Join-Path $OutputFolder "$baseName-Toevoegen.ps1"

$summaryRows = @(
    [pscustomobject]@{ Onderdeel = 'AD-groep'; Waarde = [string]$group.Name }
    [pscustomobject]@{ Onderdeel = 'SamAccountName'; Waarde = [string]$group.SamAccountName }
    [pscustomobject]@{ Onderdeel = 'DistinguishedName'; Waarde = [string]$group.DistinguishedName }
    [pscustomobject]@{ Onderdeel = 'Beschrijving'; Waarde = [string]$group.Description }
    [pscustomobject]@{ Onderdeel = 'Geneste leden meegenomen'; Waarde = $(if ($Recursive) { 'Ja' } else { 'Nee' }) }
    [pscustomobject]@{ Onderdeel = 'Domeincontroller'; Waarde = $(if ($Server) { $Server } else { '(standaard)' }) }
    [pscustomobject]@{ Onderdeel = 'Invoerbestand'; Waarde = $InputFile }
    [pscustomobject]@{ Onderdeel = 'Gegenereerd op'; Waarde = $script:GeneratedAt.ToString('dd-MM-yyyy HH:mm') }
    [pscustomobject]@{ Onderdeel = 'Leden in groep (gebruikers)'; Waarde = $groupUsers.Count }
    [pscustomobject]@{ Onderdeel = 'Verwacht aantal leden'; Waarde = $(if ($script:ExpectedCountOpgegeven) { $ExpectedCount } else { '(niet opgegeven)' }) }
    [pscustomobject]@{ Onderdeel = 'Unieke e-mailadressen op lijst'; Waarde = $emails.Count }
    [pscustomobject]@{ Onderdeel = 'Overeenkomend'; Waarde = $matchedRows.Count }
    [pscustomobject]@{ Onderdeel = 'Toe te voegen (bestaan in AD)'; Waarde = $toAddRows.Count }
    [pscustomobject]@{ Onderdeel = 'Niet gevonden in AD'; Waarde = $notFoundRows.Count }
    [pscustomobject]@{ Onderdeel = 'Staat in groep, niet op lijst'; Waarde = $extraRows.Count }
    [pscustomobject]@{ Onderdeel = 'Ongeldige regels'; Waarde = @($invalidEntries).Count }
    [pscustomobject]@{ Onderdeel = 'Dubbele regels'; Waarde = @($duplicateEntries).Count }
    [pscustomobject]@{ Onderdeel = 'Niet-gebruikers overgeslagen'; Waarde = $skippedNonUsers }
    [pscustomobject]@{ Onderdeel = 'Leden die niet uitgelezen konden worden'; Waarde = $failedMembers }
)

# --- Rapport wegschrijven -----------------------------------------------------
$useExcel = Test-ImportExcelAvailable
if (-not $useExcel) {
    $useExcel = Initialize-ImportExcelModule
}
elseif (-not (Get-Module -Name ImportExcel)) {
    try {
        Import-Module ImportExcel -ErrorAction Stop
    }
    catch {
        Write-Status "ImportExcel kon niet worden geladen: $($_.Exception.Message). Er worden CSV-bestanden gemaakt." 'Yellow'
        $useExcel = $false
    }
}

$exportedFiles = [System.Collections.Generic.List[string]]::new()

if ($useExcel) {
    try {
        if (Test-Path -LiteralPath $excelPath) {
            Remove-Item -LiteralPath $excelPath -Force
        }

        $excel = $summaryRows | Export-Excel -Path $excelPath -WorksheetName 'Samenvatting' `
            -AutoSize -BoldTopRow -FreezeTopRow -TableName 'Samenvatting' -TableStyle 'Medium2' -PassThru
        try {
            $ws = $excel.Workbook.Worksheets['Samenvatting']
            if ($ws) {
                $ws.Cells['A1:B1'].Style.Font.Bold = $true
                $ws.Column(1).Width = 46
                $ws.Column(2).Width = 70
            }
            $excel.Save()
        }
        finally {
            $excel.Dispose()
        }

        Add-WorksheetIfData -ExcelPath $excelPath -WorksheetName 'Toevoegen aan groep' -Data $toAddRows -TableStyle 'Medium6'
        Add-WorksheetIfData -ExcelPath $excelPath -WorksheetName 'Niet op lijst' -Data $extraRows -TableStyle 'Medium4'
        Add-WorksheetIfData -ExcelPath $excelPath -WorksheetName 'Overeenkomend' -Data $matchedRows -TableStyle 'Medium7'
        Add-WorksheetIfData -ExcelPath $excelPath -WorksheetName 'Niet gevonden in AD' -Data $notFoundRows -TableStyle 'Medium3'
        if ($invalidRows.Count -gt 0) {
            Add-WorksheetIfData -ExcelPath $excelPath -WorksheetName 'Ongeldige regels' -Data $invalidRows -TableStyle 'Medium9'
        }

        $exportedFiles.Add($excelPath)
        Write-Status "Excel-rapport opgeslagen: $excelPath" 'Green'
    }
    catch {
        Write-Status "Excel-rapport opslaan mislukt: $($_.Exception.Message). Er wordt teruggevallen op CSV." 'Yellow'
        $useExcel = $false
    }
}

if (-not $useExcel) {
    $csvSummary = Join-Path $OutputFolder "$baseName-Samenvatting.csv"
    $csvAdd     = Join-Path $OutputFolder "$baseName-Toevoegen.csv"
    $csvExtra   = Join-Path $OutputFolder "$baseName-NietOpLijst.csv"
    $csvMatch   = Join-Path $OutputFolder "$baseName-Overeenkomend.csv"
    $csvMissing = Join-Path $OutputFolder "$baseName-NietGevonden.csv"
    $csvInvalid = Join-Path $OutputFolder "$baseName-OngeldigeRegels.csv"
    try {
        Export-CsvSafe -Path $csvSummary -Data $summaryRows
        $exportedFiles.Add($csvSummary)
        Export-CsvSafe -Path $csvAdd -Data $toAddRows
        if ($toAddRows.Count -gt 0) { $exportedFiles.Add($csvAdd) }
        Export-CsvSafe -Path $csvExtra -Data $extraRows
        if ($extraRows.Count -gt 0) { $exportedFiles.Add($csvExtra) }
        Export-CsvSafe -Path $csvMatch -Data $matchedRows
        if ($matchedRows.Count -gt 0) { $exportedFiles.Add($csvMatch) }
        Export-CsvSafe -Path $csvMissing -Data $notFoundRows
        if ($notFoundRows.Count -gt 0) { $exportedFiles.Add($csvMissing) }
        if ($invalidRows.Count -gt 0) {
            Export-CsvSafe -Path $csvInvalid -Data $invalidRows
            $exportedFiles.Add($csvInvalid)
        }
        Write-Status "CSV-rapport(en) opgeslagen in $OutputFolder" 'Green'
    }
    catch {
        Write-Status "CSV-rapport opslaan mislukt: $($_.Exception.Message)" 'Red'
        return
    }
}

if (-not $SkipCommandFile -and $toAddUsers.Count -gt 0) {
    $commandLines = [System.Collections.Generic.List[string]]::new()
    $commandLines.Add('# Kant-en-klare commando''s om ontbrekende gebruikers toe te voegen.')
    $commandLines.Add('# Dit bestand is alleen ter controle: het vergelijkingsscript heeft niets gewijzigd.')
    $commandLines.Add("# Groep : $($group.SamAccountName)")
    $commandLines.Add("# Gegenereerd op: $($script:GeneratedAt.ToString('dd-MM-yyyy HH:mm'))")
    $commandLines.Add('# Controleer de lijst voordat je deze regels uitvoert.')
    $commandLines.Add('')
    $commandLines.Add("Import-Module ActiveDirectory")
    $safeGroup = $group.SamAccountName.Replace("'", "''")
    foreach ($user in $toAddUsers) {
        $safeSam = $user.SamAccountName.Replace("'", "''")
        $label = if ($user.LijstAdres) { $user.LijstAdres } else { $user.Mail }
        $commandLines.Add("Add-ADGroupMember -Identity '$safeGroup' -Members '$safeSam'  # $label")
    }
    try {
        $cmdEncoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
        $commandLines | Set-Content -LiteralPath $commandPath -Encoding $cmdEncoding
        $exportedFiles.Add($commandPath)
        Write-Status "Commando-lijst (niet uitgevoerd) opgeslagen: $commandPath" 'Green'
    }
    catch {
        Write-Status "Commando-lijst opslaan mislukt: $($_.Exception.Message)" 'Yellow'
    }
}

# --- Samenvatting -------------------------------------------------------------
Write-Host ''
Write-Host '================ SAMENVATTING ================' -ForegroundColor White
Write-Host ("  {0,-40} {1}" -f 'AD-groep', $group.SamAccountName) -ForegroundColor White
Write-Host ("  {0,-40} {1,5}" -f 'Leden in groep (gebruikers)', $groupUsers.Count) -ForegroundColor White
Write-Host ("  {0,-40} {1,5}" -f 'Unieke e-mailadressen op lijst', $emails.Count) -ForegroundColor White
Write-Host ("  {0,-40} {1,5}" -f 'Overeenkomend', $matchedRows.Count) -ForegroundColor Green
Write-Host ("  {0,-40} {1,5}" -f 'Toe te voegen (bestaan in AD)', $toAddRows.Count) -ForegroundColor Yellow
Write-Host ("  {0,-40} {1,5}" -f 'Niet gevonden in AD', $notFoundRows.Count) -ForegroundColor Red
Write-Host ("  {0,-40} {1,5}" -f 'Staat in groep, niet op lijst', $extraRows.Count) -ForegroundColor Cyan
if ($invalidEntries) {
    Write-Host ("  {0,-40} {1,5}" -f 'Ongeldige regels', @($invalidEntries).Count) -ForegroundColor DarkYellow
}
if ($duplicateEntries) {
    Write-Host ("  {0,-40} {1,5}" -f 'Dubbele regels', @($duplicateEntries).Count) -ForegroundColor DarkYellow
}
if ($script:ExpectedCountOpgegeven -and $groupUsers.Count -ne $ExpectedCount) {
    Write-Host ("  {0,-40} {1,5}" -f 'Afwijking t.o.v. verwacht aantal', ($groupUsers.Count - $ExpectedCount)) -ForegroundColor Yellow
}
Write-Host '==============================================' -ForegroundColor White

$primaryFile = if ($useExcel -and (Test-Path -LiteralPath $excelPath)) { $excelPath } else { $exportedFiles | Select-Object -First 1 }
if ($primaryFile) {
    Write-Status "Rapport opgeslagen: $primaryFile" 'Green'
    $answer = Read-Host 'Rapport nu openen? (J/N)'
    if ($answer -in 'J', 'j', 'Ja', 'ja') {
        Invoke-Item -LiteralPath $primaryFile
    }
}
