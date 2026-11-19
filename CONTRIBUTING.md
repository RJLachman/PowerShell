# Bijdragen

Bedankt voor je interesse! Houd bij wijzigingen de volgende afspraken aan.

## Codestijl
- Gebruik goedgekeurde werkwoorden (`Get-Verb`) en **Verb-Noun** naamgeving in PascalCase.
- Elk script heeft volledige comment-based help (`.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER`, `.EXAMPLE`, `.NOTES`).
- Gebruik `[CmdletBinding()]` en parametervalidatie (`[ValidateNotNullOrEmpty()]`, `[ValidateRange()]`, ...).
- Scripts die iets wijzigen of verwijderen ondersteunen `-WhatIf`/`-Confirm` via `SupportsShouldProcess`.
- Interactieve bevestiging gebruikt **J/N** (accepteert `J`, `j`, `Ja`, `ja`).
- Gebruikersberichten zijn in het Nederlands.
- **Nooit** secrets of organisatiespecifieke waarden (domeinen, OU's, e-mailadressen) committen; gebruik voorbeeldwaarden zoals `contoso.nl`.
- Inspringen met 4 spaties, bestanden in UTF-8 met BOM (compatibel met Windows PowerShell 5.1).

## PSScriptAnalyzer
Alle code moet zonder waarschuwingen door PSScriptAnalyzer komen met de instellingen uit de repo:

```powershell
Install-Module PSScriptAnalyzer -Scope CurrentUser
Invoke-ScriptAnalyzer -Path . -Recurse -Settings ./PSScriptAnalyzerSettings.psd1
```

Deze controle draait ook automatisch via GitHub Actions bij elke push en pull request.

## Commits
Gebruik [Conventional Commits](https://www.conventionalcommits.org/nl/v1.0.0/):

```
feat(reporting): voeg export naar Excel toe
fix(entraid): corrigeer filter op laatste aanmelding
docs: werk README bij
refactor(exchange): splits validatie in aparte functie
ci: update lint-workflow
```

Toegestane types: `feat`, `fix`, `docs`, `refactor`, `perf`, `test`, `ci`, `chore`, `style`.

## Pull requests
1. Maak een branch vanaf `main` (bijv. `feat/korte-omschrijving`).
2. Zorg dat PSScriptAnalyzer groen is.
3. Werk de README van de map en de [CHANGELOG](CHANGELOG.md) (sectie *Unreleased*) bij.
4. Open een pull request en vul het template in.
