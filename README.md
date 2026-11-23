# PowerShell

[![Lint](https://github.com/RJLachman/PowerShell/actions/workflows/lint.yml/badge.svg?branch=main)](https://github.com/RJLachman/PowerShell/actions/workflows/lint.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207%2B-5391FE?logo=powershell&logoColor=white)](https://learn.microsoft.com/powershell/)

Een verzameling praktische PowerShell-scripts voor het beheer van Active Directory, Exchange (hybrid) en Microsoft 365. Alle scripts hebben comment-based help (`Get-Help .\Script.ps1 -Full`), Nederlandse gebruikersberichten en worden bij elke push automatisch gecontroleerd met PSScriptAnalyzer.

## Scripts

| Map | Script | Omschrijving |
|---|---|---|
| [ActiveDirectory](ActiveDirectory) | [Compare-ADGroupMembership.ps1](ActiveDirectory/Compare-ADGroupMembership.ps1) | Vergelijkt de leden van een AD-groep met een e-maillijst en maakt een Excel-rapport. Alleen-lezen. |
| [ActiveDirectory](ActiveDirectory) | [Copy-ADGroupMembership.ps1](ActiveDirectory/Copy-ADGroupMembership.ps1) | Kopieert groepslidmaatschappen van een brongebruiker naar een doelgebruiker (op e-mailadres). |
| [Exchange](Exchange) | [New-SharedRemoteMailbox.ps1](Exchange/New-SharedRemoteMailbox.ps1) | Maakt een shared remote mailbox aan (hybrid) en zet de AD Description. |
| [Microsoft365/EntraID](Microsoft365/EntraID) | [Remove-StaleAzureADDevices.ps1](Microsoft365/EntraID/Remove-StaleAzureADDevices.ps1) | Exporteert en verwijdert Entra ID-apparaten die al X dagen niet zijn aangemeld (Microsoft Graph). |
| [Microsoft365/Purview](Microsoft365/Purview) | [Remove-EmailFromSender.ps1](Microsoft365/Purview/Remove-EmailFromSender.ps1) | Zoekt via Content Search alle mails van een afzender en soft-delete ze na bevestiging. |
| [Reporting](Reporting) | [Get-UserLastLogonReport.ps1](Reporting/Get-UserLastLogonReport.ps1) | CSV-rapport van de laatste aanmelding van gebruikers (AD en/of Entra ID). Alleen-lezen. |

Elke map heeft een eigen README met details, vereisten en voorbeelden.

Scripts die iets wijzigen of verwijderen ondersteunen **`-WhatIf`** (laat zien wat er zou gebeuren), **`-Confirm`** en **`-Force`** (slaat de J/N-vraag over). Zonder deze parameters vragen ze altijd eerst om een J/N-bevestiging (`J`, `j`, `Ja` of `ja`).

## Vereisten

- Windows PowerShell 5.1 of PowerShell 7+
- Per script de benodigde module(s), bijvoorbeeld:
  - `ActiveDirectory` (RSAT) voor de AD- en rapportagescripts
  - `ImportExcel` (optioneel) voor het Excel-rapport van Compare-ADGroupMembership
  - Exchange Management Shell voor de Exchange-scripts
  - `Microsoft.Graph` voor Entra ID
  - `ExchangeOnlineManagement` voor Purview (`Connect-IPPSSession`)
- Voldoende beheerrechten in de betreffende omgeving

## Gebruik

```powershell
# Repository klonen
git clone https://github.com/RJLachman/PowerShell.git
cd PowerShell

# Help van een script bekijken
Get-Help .\ActiveDirectory\Copy-ADGroupMembership.ps1 -Full

# Eerst droog draaien: laat zien wat er zou gebeuren, zonder iets te wijzigen
.\ActiveDirectory\Copy-ADGroupMembership.ps1 -WhatIf

# Groepsleden vergelijken met een e-maillijst (alleen-lezen, Excel-rapport)
.\ActiveDirectory\Compare-ADGroupMembership.ps1 -GroupName 'soft-apps-154' -InputFile 'C:\temp\suwi.txt' -ExpectedCount 84
```

Gedownloade scripts zijn mogelijk geblokkeerd door Windows. Deblokkeer ze met `Unblock-File` of stel een passende execution policy in (bijv. `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`).

Voorbeeldwaarden zoals `contoso.nl`, `contoso.local` en `afzender@voorbeeld.nl` vervang je door de waarden van je eigen omgeving (meestal via een parameter).

## Disclaimer

> ⚠️ **Gebruik op eigen risico; test eerst met `-WhatIf`.**
> Een aantal scripts wijzigt of verwijdert gegevens (groepslidmaatschappen, mailboxen, apparaten, e-mail). Test altijd eerst in een testomgeving of met `-WhatIf` en controleer de uitvoer voordat je bevestigt.

## Bijdragen en beveiliging

- Zie [CONTRIBUTING.md](CONTRIBUTING.md) voor codestijl, PSScriptAnalyzer en commit-afspraken.
- Beveiligingsproblemen meld je privé, zie [SECURITY.md](SECURITY.md).
- Wijzigingen per versie staan in de [CHANGELOG](CHANGELOG.md).

## Licentie

Uitgebracht onder de [MIT-licentie](LICENSE). © 2026 Roy James Lachman.
