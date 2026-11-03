<#
.SYNOPSIS
 Maakt een nieuwe gedeelde (shared) remote mailbox aan voor contoso.nl en zet de AD Description.

.DESCRIPTION
 - Vraagt de weergavenaam (Display Name).
 - Genereert automatisch een alias (mailNickname) en UPN.
 - Controleert of e-mailadres/alias al bestaat.
 - Maakt de shared remote mailbox aan in de opgegeven OU.
 - Stelt de AD Description in op: "Shared Mailbox - <DisplayName>".

.PARAMETER Domain
 Het domein (standaard: contoso.nl).

.PARAMETER OU
 De on-prem OU waarin de mailbox wordt aangemaakt 
 (standaard: contoso.local/Shared Mailboxes).

.NOTES
 Auteur   : rjlachman 😎
 Versie   : 1.2
 Laatst gewijzigd : 8 oktober 2025
 Vereisten: Exchange Management Shell (on-prem); optioneel RSAT AD module
#>

[CmdletBinding()]
param(
    [Parameter()][ValidateNotNullOrEmpty()]
    [string]$Domain = "contoso.nl",

    [Parameter()][ValidateNotNullOrEmpty()]
    [string]$OU = "contoso.local/Shared Mailboxes"
)

function New-SafeAlias {
<#
.SYNOPSIS
 Zet een display name om naar een geldig alias.
#>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DisplayName)
    
    $alias = $DisplayName.ToLowerInvariant() -replace '[^a-z0-9]+','.'
    $alias = ($alias -replace '\.+','.') -replace '^\.+|\.+$',''
    if ([string]::IsNullOrWhiteSpace($alias)) {
        $alias = "shared.mailbox.$([DateTime]::UtcNow.ToString('yyyyMMddHHmmss'))"
    }
    return $alias
}

function Test-RecipientExists {
<#
.SYNOPSIS
 Controleert of een e-mailadres of alias al bestaat.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SmtpAddress,
        [Parameter(Mandatory)][string]$Alias
    )

    $r1 = Get-Recipient -Identity $SmtpAddress -ErrorAction SilentlyContinue
    if ($r1) { return $true }

    $filter = "EmailAddresses -eq 'SMTP:$SmtpAddress' -or EmailAddresses -eq 'smtp:$SmtpAddress'"
    $r2 = Get-Recipient -ResultSize Unlimited -Filter $filter -ErrorAction SilentlyContinue
    if ($r2) { return $true }

    $r3 = Get-Recipient -ResultSize Unlimited -Filter "Alias -eq '$Alias'" -ErrorAction SilentlyContinue
    if ($r3) { return $true }

    return $false
}

function Set-ADDescriptionForMailbox {
<#
.SYNOPSIS
 Zet de AD Description voor de zojuist aangemaakte (remote) mailbox.
.DESCRIPTION
 Probeert eerst via de ActiveDirectory module (Set-ADUser). 
 Als die niet beschikbaar is, gebruikt hij ADSI met de DistinguishedName uit Get-User.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$IdentityUpn,
        [Parameter(Mandatory)][string]$DescriptionText
    )

    $usingAdModule = $false
    try {
        if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) { throw "AD module niet gevonden." }
        Import-Module ActiveDirectory -ErrorAction Stop
        $usingAdModule = $true
    } catch {
        # Geen AD module, we vallen later terug op ADSI
    }

    if ($usingAdModule) {
        try {
            $adUser = Get-ADUser -Filter "userPrincipalName -eq '$IdentityUpn'" -Properties Description -ErrorAction Stop
            Set-ADUser -Identity $adUser -Description $DescriptionText -ErrorAction Stop
            Write-Host "AD Description gezet via AD-module: '$DescriptionText'." -ForegroundColor DarkGreen
            return
        } catch {
            Write-Warning "Kon AD Description niet zetten via AD-module: $($_.Exception.Message). Probeer ADSI fallback..."
        }
    }

    # Fallback via ADSI
    try {
        $dn = (Get-User -Identity $IdentityUpn -ErrorAction Stop).DistinguishedName
        if (-not $dn) { throw "DistinguishedName niet gevonden via Get-User." }
        $de = [ADSI]"LDAP://$dn"
        $de.description = $DescriptionText
        $de.SetInfo()
        Write-Host "AD Description gezet via ADSI: '$DescriptionText'." -ForegroundColor DarkGreen
    } catch {
        Write-Warning "Kon AD Description niet zetten via ADSI fallback: $($_.Exception.Message)"
    }
}

# --- Voorchecks ---
if (-not (Get-Command New-RemoteMailbox -ErrorAction SilentlyContinue)) {
    Write-Error "De cmdlet 'New-RemoteMailbox' is niet gevonden. Start de Exchange Management Shell of laad de Exchange snapins/modules."
    exit 1
}

# --- Input ---
$displayName = Read-Host "Voer de weergavenaam in voor de nieuwe shared mailbox (bijv. 'Afdeling Inkoop')"
if ([string]::IsNullOrWhiteSpace($displayName)) {
    Write-Error "Geen weergavenaam opgegeven. Script wordt beëindigd."
    exit 1
}

$proposedAlias = New-SafeAlias -DisplayName $displayName
$alias = Read-Host "Alias/e-mailadres (zonder @${Domain}) [Enter = $proposedAlias]"
if ([string]::IsNullOrWhiteSpace($alias)) {
    $alias = $proposedAlias
} else {
    $alias = $alias.ToLowerInvariant()
    $alias = ($alias -replace '[^a-z0-9]+','.') -replace '\.+','.'
    $alias = $alias -replace '^\.+|\.+$',''
    if ([string]::IsNullOrWhiteSpace($alias)) {
        Write-Error "Ongeldige alias na schoonmaken."
        exit 1
    }
}

# Afgeleide waarden
$upn = "$alias@$Domain"
$sam = if ($alias.Length -gt 20) { $alias.Substring(0,20) } else { $alias }
$adDescription = "Shared Mailbox - $displayName"

Write-Host ""
Write-Host "Overzicht:" -ForegroundColor Cyan
Write-Host "  Display Name     : $displayName"
Write-Host "  Alias            : $alias"
Write-Host "  UPN / PrimarySMTP: $upn"
Write-Host "  SAM-accountnaam  : $sam"
Write-Host "  OU               : $OU"
Write-Host "  AD Description   : $adDescription"
Write-Host ""

# --- Controle of adres/alias al bestaat ---
if (Test-RecipientExists -SmtpAddress $upn -Alias $alias) {
    Write-Warning "Er bestaat al een object met e-mailadres of alias '$upn' / '$alias'. Kies een andere alias/naam."
    exit 1
}

# --- Bevestiging ---
$go = Read-Host "Shared mailbox aanmaken met bovenstaande gegevens? (J/N)"
if ($go -notin @('J','j','Y','y','Ja','ja')) {
    Write-Host "Geannuleerd. Er is niets gewijzigd."
    exit 0
}

# --- Aanmaken ---
try {
    New-RemoteMailbox `
        -Shared `
        -Name $displayName `
        -DisplayName $displayName `
        -Alias $alias `
        -UserPrincipalName $upn `
        -PrimarySmtpAddress $upn `
        -SamAccountName $sam `
        -OnPremisesOrganizationalUnit $OU `
        -ErrorAction Stop

    Write-Host ""
    Write-Host "✅ Shared remote mailbox is aangemaakt: $displayName <$upn>" -ForegroundColor Green

    # AD Description zetten
    Set-ADDescriptionForMailbox -IdentityUpn $upn -DescriptionText $adDescription
    Write-Host "Klaar. (Controleer hybrid routing en laat AAD Connect syncen.)"
}
catch {
    Write-Error "Er ging iets mis bij het aanmaken: $($_.Exception.Message)"
    exit 1
}
