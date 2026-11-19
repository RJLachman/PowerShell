<#
.SYNOPSIS
    Maakt een nieuwe gedeelde (shared) remote mailbox aan en zet de AD Description.

.DESCRIPTION
    Bedoeld voor Exchange on-premises in een hybride omgeving met Exchange Online.
    - Vraagt de weergavenaam (Display Name).
    - Genereert automatisch een alias (mailNickname), UPN en SAM-accountnaam.
    - Controleert of het e-mailadres of de alias al bestaat.
    - Maakt de shared remote mailbox aan in de opgegeven OU (New-RemoteMailbox -Shared).
    - Stelt de AD Description in op: "Shared Mailbox - <DisplayName>".

    Met -WhatIf zie je wat er aangemaakt zou worden, zonder iets te wijzigen.
    Met -Force wordt de J/N-vraag overgeslagen.

.PARAMETER Domain
    Het e-maildomein voor de UPN en het primaire SMTP-adres (standaard: contoso.nl).

.PARAMETER OU
    De on-premises OU waarin de mailbox wordt aangemaakt (standaard: contoso.local/Shared Mailboxes).

.PARAMETER DisplayName
    Weergavenaam van de nieuwe shared mailbox. Wordt gevraagd als je het weglaat.

.PARAMETER Alias
    Alias (deel van het e-mailadres vóór de @). Wordt voorgesteld op basis van de weergavenaam
    en gevraagd als je het weglaat.

.PARAMETER Force
    Slaat de J/N-bevestiging over.

.EXAMPLE
    .\New-SharedRemoteMailbox.ps1 -Domain 'contoso.nl' -OU 'contoso.local/Shared Mailboxes'

    Vraagt om weergavenaam en alias, toont een overzicht en vraagt (J/N) om bevestiging.

.EXAMPLE
    .\New-SharedRemoteMailbox.ps1 -DisplayName 'Afdeling Inkoop' -WhatIf

    Toont welke mailbox aangemaakt zou worden (alias wordt gevraagd), zonder iets te wijzigen.

.EXAMPLE
    .\New-SharedRemoteMailbox.ps1 -DisplayName 'Afdeling Inkoop' -Alias 'inkoop' -Force

    Maakt inkoop@contoso.nl aan zonder J/N-vraag.

.NOTES
    Auteur   : RJ Lachman
    Versie   : 1.3.0
    Vereisten: Exchange Management Shell (on-premises); optioneel module ActiveDirectory (RSAT),
               anders wordt ADSI gebruikt voor de Description.
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$Domain = 'contoso.nl',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OU = 'contoso.local/Shared Mailboxes',

    [Parameter()]
    [string]$DisplayName,

    [Parameter()]
    [string]$Alias,

    [Parameter()]
    [switch]$Force
)

function ConvertTo-SafeAlias {
    <#
    .SYNOPSIS
        Zet een tekst om naar een geldige Exchange-alias (kleine letters, cijfers en punten).
    .PARAMETER Text
        De weergavenaam of opgegeven alias.
    .PARAMETER AllowEmpty
        Geeft een lege string terug in plaats van een gegenereerde alias als er niets overblijft.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][string]$Text,
        [switch]$AllowEmpty
    )

    $result = $Text.ToLowerInvariant() -replace '[^a-z0-9]+', '.'
    $result = ($result -replace '\.+', '.') -replace '^\.+|\.+$', ''
    if ([string]::IsNullOrWhiteSpace($result) -and -not $AllowEmpty) {
        $result = "shared.mailbox.$([DateTime]::UtcNow.ToString('yyyyMMddHHmmss'))"
    }
    return $result
}

function Test-RecipientExist {
    <#
    .SYNOPSIS
        Controleert of een e-mailadres of alias al bestaat in Exchange.
    .PARAMETER SmtpAddress
        Het te controleren SMTP-adres.
    .PARAMETER Alias
        De te controleren alias.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)][string]$SmtpAddress,
        [Parameter(Mandatory)][string]$Alias
    )

    # SilentlyContinue: 'niet gevonden' is hier de verwachte, goede uitkomst en geen fout.
    if (Get-Recipient -Identity $SmtpAddress -ErrorAction SilentlyContinue) { return $true }

    $filter = "EmailAddresses -eq 'SMTP:$SmtpAddress' -or EmailAddresses -eq 'smtp:$SmtpAddress'"
    if (Get-Recipient -ResultSize Unlimited -Filter $filter -ErrorAction SilentlyContinue) { return $true }

    if (Get-Recipient -ResultSize Unlimited -Filter "Alias -eq '$Alias'" -ErrorAction SilentlyContinue) { return $true }

    return $false
}

function Set-ADDescriptionForMailbox {
    <#
    .SYNOPSIS
        Zet de AD Description voor de zojuist aangemaakte (remote) mailbox.
    .DESCRIPTION
        Probeert eerst via de ActiveDirectory-module (Set-ADUser).
        Als die niet beschikbaar is, wordt ADSI gebruikt met de DistinguishedName uit Get-User.
    .PARAMETER IdentityUpn
        UPN van het mailboxaccount.
    .PARAMETER DescriptionText
        De in te stellen Description.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$IdentityUpn,
        [Parameter(Mandatory)][string]$DescriptionText
    )

    if (-not $PSCmdlet.ShouldProcess($IdentityUpn, "AD Description instellen op '$DescriptionText'")) { return }

    $usingAdModule = $false
    if (Get-Module -Name ActiveDirectory -ListAvailable) {
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
            $usingAdModule = $true
        }
        catch {
            Write-Verbose "ActiveDirectory-module kon niet worden geladen ($($_.Exception.Message)); ADSI wordt gebruikt."
        }
    }
    else {
        Write-Verbose 'ActiveDirectory-module niet gevonden; ADSI wordt gebruikt.'
    }

    if ($usingAdModule) {
        try {
            $adUser = Get-ADUser -Filter "userPrincipalName -eq '$IdentityUpn'" -Properties Description -ErrorAction Stop
            if (-not $adUser) { throw "Gebruiker '$IdentityUpn' niet gevonden in AD (nog niet gerepliceerd?)." }
            Set-ADUser -Identity $adUser -Description $DescriptionText -ErrorAction Stop
            Write-Host "AD Description gezet via AD-module: '$DescriptionText'." -ForegroundColor DarkGreen
            return
        }
        catch {
            Write-Warning "Kon AD Description niet zetten via AD-module: $($_.Exception.Message). Probeer ADSI fallback..."
        }
    }

    try {
        $dn = (Get-User -Identity $IdentityUpn -ErrorAction Stop).DistinguishedName
        if (-not $dn) { throw 'DistinguishedName niet gevonden via Get-User.' }
        $de = [ADSI]"LDAP://$dn"
        $de.description = $DescriptionText
        $de.SetInfo()
        Write-Host "AD Description gezet via ADSI: '$DescriptionText'." -ForegroundColor DarkGreen
    }
    catch {
        Write-Warning "Kon AD Description niet zetten via ADSI fallback: $($_.Exception.Message)"
    }
}

# --- Voorchecks ---------------------------------------------------------------
if (-not (Get-Command New-RemoteMailbox -ErrorAction SilentlyContinue)) {
    Write-Error "De cmdlet 'New-RemoteMailbox' is niet gevonden. Start de Exchange Management Shell of laad de Exchange snapins/modules."
    exit 1
}

# --- Input --------------------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($DisplayName)) {
    $DisplayName = Read-Host "Voer de weergavenaam in voor de nieuwe shared mailbox (bijv. 'Afdeling Inkoop')"
}
$DisplayName = "$DisplayName".Trim()
if ([string]::IsNullOrWhiteSpace($DisplayName)) {
    Write-Error 'Geen weergavenaam opgegeven. Script wordt beëindigd.'
    exit 1
}

$proposedAlias = ConvertTo-SafeAlias -Text $DisplayName
if ([string]::IsNullOrWhiteSpace($Alias)) {
    $Alias = Read-Host "Alias/e-mailadres (zonder @${Domain}) [Enter = $proposedAlias]"
}
if ([string]::IsNullOrWhiteSpace($Alias)) {
    $Alias = $proposedAlias
}
else {
    $Alias = ConvertTo-SafeAlias -Text $Alias -AllowEmpty
    if ([string]::IsNullOrWhiteSpace($Alias)) {
        Write-Error 'Ongeldige alias na schoonmaken.'
        exit 1
    }
}

# Afgeleide waarden
$upn = "$Alias@$Domain"
$sam = if ($Alias.Length -gt 20) { $Alias.Substring(0, 20) } else { $Alias }
$adDescription = "Shared Mailbox - $DisplayName"

Write-Host ''
Write-Host 'Overzicht:' -ForegroundColor Cyan
Write-Host "  Display Name     : $DisplayName"
Write-Host "  Alias            : $Alias"
Write-Host "  UPN / PrimarySMTP: $upn"
Write-Host "  SAM-accountnaam  : $sam"
Write-Host "  OU               : $OU"
Write-Host "  AD Description   : $adDescription"
Write-Host ''

# --- Controle of adres/alias al bestaat ---------------------------------------
if (Test-RecipientExist -SmtpAddress $upn -Alias $Alias) {
    Write-Warning "Er bestaat al een object met e-mailadres of alias '$upn' / '$Alias'. Kies een andere alias/naam."
    exit 1
}

# --- Bevestiging --------------------------------------------------------------
if (-not $WhatIfPreference -and -not $Force) {
    $go = Read-Host 'Shared mailbox aanmaken met bovenstaande gegevens? (J/N)'
    if ($go -notin @('J', 'j', 'Y', 'y', 'Ja', 'ja')) {
        Write-Host 'Geannuleerd. Er is niets gewijzigd.'
        exit 0
    }
}
# De J/N-vraag (of -Force) is de bevestiging; geen tweede prompt tenzij -Confirm is opgegeven.
if (-not $PSBoundParameters.ContainsKey('Confirm')) { $ConfirmPreference = 'None' }

# --- Aanmaken -----------------------------------------------------------------
if (-not $PSCmdlet.ShouldProcess("$DisplayName <$upn> in OU '$OU'", 'Shared remote mailbox aanmaken')) {
    exit 0
}

try {
    $mailboxParams = @{
        Shared                       = $true
        Name                         = $DisplayName
        DisplayName                  = $DisplayName
        Alias                        = $Alias
        UserPrincipalName            = $upn
        PrimarySmtpAddress           = $upn
        SamAccountName               = $sam
        OnPremisesOrganizationalUnit = $OU
        ErrorAction                  = 'Stop'
    }
    New-RemoteMailbox @mailboxParams

    Write-Host ''
    Write-Host "✅ Shared remote mailbox is aangemaakt: $DisplayName <$upn>" -ForegroundColor Green

    Set-ADDescriptionForMailbox -IdentityUpn $upn -DescriptionText $adDescription
    Write-Host 'Klaar. (Controleer hybrid routing en laat Entra Connect syncen.)'
}
catch {
    Write-Error "Er ging iets mis bij het aanmaken: $($_.Exception.Message)"
    exit 1
}
