# Reporting

## Get-UserLastLogonReport.ps1

Maakt een overzichtelijk CSV-rapport van wanneer gebruikers voor het laatst zijn aangemeld.

**Werking**
1. Leest e-mailadressen uit `C:\temp\userslogin.txt` (één per regel; lege regels en `#`-regels worden genegeerd, dubbele adressen worden ontdubbeld).
2. Zoekt elke gebruiker op in **on-prem AD** (`LastLogonTimestamp`, standaard) en optioneel in **Entra ID** (Microsoft Graph `signInActivity`, via `-Source EntraID` of `-Source Both`).
3. Exporteert naar `C:\temp\UserLastLogonReport_<datum>.csv` met `;` als scheidingsteken (opent direct goed in Nederlandse Excel).
4. Toont een gekleurde samenvatting en vraagt (J/N) of het rapport geopend moet worden.

**Kolommen:** Weergavenaam, Voornaam, Achternaam, E-mailadres, UPN, Afdeling, Functie, Account actief, Status, Laatste aanmelding, Dagen geleden, Laatste aanmelding Entra, Laatste aanmelding AD, Aangemaakt op, Gevonden in, Opmerkingen.

**Status:** `Actief`, `Inactief (> X dagen)`, `Nooit aangemeld`, `Uitgeschakeld`, `Niet gevonden`. Het rapport is gesorteerd met de langst inactieve gebruikers bovenaan.

**Vereisten**
- Entra ID: module `Microsoft.Graph` (`Install-Module Microsoft.Graph -Scope CurrentUser`), scopes `User.Read.All` en `AuditLog.Read.All`, en een Entra ID P1/P2-licentie voor `signInActivity`.
- AD: RSAT-module `ActiveDirectory`.

**Voorbeelden**
```powershell
.\Get-UserLastLogonReport.ps1                                   # lokale AD, standaardpaden
.\Get-UserLastLogonReport.ps1 -Source Both                      # AD + Entra ID
.\Get-UserLastLogonReport.ps1 -Source EntraID -InactiveDays 60  # alleen Entra ID, inactief na 60 dagen
.\Get-UserLastLogonReport.ps1 -ExactADLogon                     # exacte AD-logon via alle DC's (trager)
```

> `LastLogonTimestamp` in AD kan 9 tot 14 dagen achterlopen. Gebruik `-ExactADLogon` voor de exacte waarde.
