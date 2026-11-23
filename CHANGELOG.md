# Changelog

Alle noemenswaardige wijzigingen in dit project worden in dit bestand bijgehouden.

Het formaat is gebaseerd op [Keep a Changelog](https://keepachangelog.com/nl/1.1.0/)
en dit project volgt [Semantic Versioning](https://semver.org/lang/nl/).

## [Unreleased]

## [1.3.0] - 2026-11-23

### Gewijzigd
- **ActiveDirectory/Compare-UserListToADGroup.ps1** (voorheen **Compare-ADGroupMembership.ps1**) – hernoemd, zodat de bestandsnaam beter weergeeft dat het script een lijst gebruikers (e-mailadressen) controleert tegen een AD-groep. Aanroep, parameters en rapportage zijn verder gelijk.

## [1.2.0] - 2026-11-23

### Gewijzigd
- **Compare-ADGroupMembership.ps1** – geen vaste groepsnaam of invoerbestand meer. Zonder `-GroupName` kun je op naam zoeken in AD en een groep kiezen; zonder `-InputFile` verschijnt een bestandsvenster (of een padvraag). Met beide parameters blijft het script niet-interactief. Voorbeelden gebruiken fictieve namen zoals `GRP-App-Voorbeeld` en `C:\temp\gebruikers.txt`.

## [1.1.1] - 2026-11-23

### Opgelost
- **Compare-ADGroupMembership.ps1** – geen waarschuwingen meer van Export-Excel over ongeldige tekens in bereiknamen (`E-mail op lijst`, `E-mail (mail)`, `Account actief`). De kolomkoppen blijven ongewijzigd.

## [1.1.0] - 2026-11-23

### Toegevoegd
- **ActiveDirectory/Compare-ADGroupMembership.ps1** – vergelijkt de gebruikersleden van een on-premises AD-groep met een e-maillijst en maakt een overzichtelijk Excel-rapport (of CSV als de module ImportExcel ontbreekt). Alleen-lezen; schrijft optioneel kant-en-klare `Add-ADGroupMember`-commando's weg zonder ze uit te voeren.

## [1.0.0] - 2026-11-19

Eerste officiële release.

### Scripts
- **ActiveDirectory/Copy-ADGroupMembership.ps1** – kopieert groepslidmaatschappen van een bron- naar een doelgebruiker (op e-mailadres).
- **Exchange/New-SharedRemoteMailbox.ps1** – maakt een shared remote mailbox aan (hybrid) en zet de AD Description.
- **Microsoft365/EntraID/Remove-StaleAzureADDevices.ps1** – exporteert en verwijdert Entra ID-apparaten die al X dagen niet zijn aangemeld (Microsoft Graph, zonder WAM, met device code fallback).
- **Microsoft365/Purview/Remove-EmailFromSender.ps1** – zoekt via Content Search mails van een afzender en soft-delete ze na bevestiging.
- **Reporting/Get-UserLastLogonReport.ps1** – CSV-rapport van de laatste aanmelding van gebruikers uit AD en/of Entra ID.

### Toegevoegd
- `LICENSE` (MIT), `.gitignore`, `SECURITY.md`, `CONTRIBUTING.md` en dit `CHANGELOG.md`.
- `PSScriptAnalyzerSettings.psd1` met gedeelde lintregels.
- GitHub Actions-workflow `lint.yml` (parse check + PSScriptAnalyzer bij push en pull request).
- Dependabot voor GitHub Actions, issue-templates en een pull request-template.
- VS Code-instellingen (`.vscode/settings.json`) die de PSScriptAnalyzer-instellingen gebruiken.
- `-WhatIf`, `-Confirm` en `-Force` voor alle scripts die iets wijzigen of verwijderen (de J/N-vraag blijft standaard).
- Nieuwe optionele parameters: `-SenderAddress` en `-PollIntervalSeconds` (Purview), `-ExportPath` (Entra ID), `-SourceEmail`/`-TargetEmail` (AD), `-DisplayName`/`-Alias` (Exchange).
- Volledige comment-based help (`Get-Help -Full`) voor alle scripts.

### Gewijzigd
- `SoftDelete-EmailFromSender.ps1` hernoemd naar **`Remove-EmailFromSender.ps1`** (goedgekeurd werkwoord).
- Root-README herschreven met badges, scriptoverzicht, vereisten, gebruik en disclaimer; map-READMEs uitgebreid met parameters.
- Alle scripts slaan op als UTF-8 met BOM (correcte weergave van é/ë/✅ in Windows PowerShell 5.1).
- Betere foutafhandeling (try/catch) rond verbinden, ophalen, exporteren en wijzigen.

### Opgelost
- Purview: het script wachtte eindeloos als een search of purge de status `Failed` kreeg; het stopt nu met een melding.
- Purview: gebruikte de automatische variabele `$sender`.
- Active Directory: enkele quotes in een e-mailadres konden het AD-filter breken.

[Unreleased]: https://github.com/RJLachman/PowerShell/compare/v1.3.0...HEAD
[1.3.0]: https://github.com/RJLachman/PowerShell/compare/v1.2.0...v1.3.0
[1.2.0]: https://github.com/RJLachman/PowerShell/compare/v1.1.1...v1.2.0
[1.1.1]: https://github.com/RJLachman/PowerShell/compare/v1.1.0...v1.1.1
[1.1.0]: https://github.com/RJLachman/PowerShell/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/RJLachman/PowerShell/releases/tag/v1.0.0
