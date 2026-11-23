# Active Directory (on-premises)

Scripts voor lokale Active Directory.

| Script | Omschrijving |
|---|---|
| [Compare-ADGroupMembership.ps1](Compare-ADGroupMembership.ps1) | Vergelijkt de gebruikersleden van een AD-groep met een lijst e-mailadressen en maakt een Excel-rapport. Alleen-lezen. |
| [Copy-ADGroupMembership.ps1](Copy-ADGroupMembership.ps1) | Kopieert de groepslidmaatschappen van een brongebruiker naar een doelgebruiker, op basis van e-mailadres. Toont eerst welke groepen erbij komen en vraagt om bevestiging. |

## Compare-ADGroupMembership.ps1

Leest e-mailadressen uit een tekstbestand en vergelijkt die met de gebruikers in een AD-groep. Het rapport laat zien welke adressen nog toegevoegd moeten worden, welke nergens in AD staan, en welke groepsleden niet op de lijst voorkomen. Een adres telt als overeenkomst als het gelijk is aan `mail`, `UserPrincipalName` of een `smtp:`-proxyAddress (hoofdletterongevoelig).

Het script wijzigt niets in Active Directory. Optioneel wordt een tekstbestand met kant-en-klare `Add-ADGroupMember`-commando's weggeschreven; die commando's worden niet uitgevoerd.

```powershell
.\Compare-ADGroupMembership.ps1
.\Compare-ADGroupMembership.ps1 -GroupName 'soft-apps-154' -InputFile 'C:\temp\suwi.txt' -ExpectedCount 84
.\Compare-ADGroupMembership.ps1 -GroupName 'APP-Finance' -InputFile 'D:\lijsten\finance.txt' -OutputFolder 'D:\rapporten' -Recursive -Server 'DC01.contoso.local'
```

| Parameter | Omschrijving |
|---|---|
| `-GroupName` | Naam van de AD-groep (standaard `soft-apps-154`). |
| `-InputFile` | Tekstbestand met e-mailadressen (standaard `C:\temp\suwi.txt`). Lege regels, commentaarregels (`#`), duplicaten en ongeldige adressen worden netjes afgehandeld. |
| `-OutputFolder` | Map voor het rapport (standaard `C:\temp`). Bestandsnaam: `ADGroepVergelijking_<groep>_yyyyMMdd_HHmm.xlsx`. |
| `-Recursive` | Neemt ook leden van geneste groepen mee. |
| `-Server` | Optionele domeincontroller voor de AD-query's. |
| `-ExpectedCount` | Optioneel verwacht aantal gebruikersleden; wijkt het af, dan volgt een waarschuwing. |
| `-SkipCommandFile` | Schrijft geen tekstbestand met `Add-ADGroupMember`-commando's weg. |

Het Excel-rapport (module `ImportExcel`) bevat de werkbladen **Samenvatting**, **Toevoegen aan groep**, **Niet op lijst**, **Overeenkomend**, **Niet gevonden in AD** en eventueel **Ongeldige regels**. Ontbreekt ImportExcel, dan biedt het script een J/N-vraag om de module te installeren (`J`, `j`, `Ja` of `ja`). Bij weigering vallen de resultaten terug op aparte CSV-bestanden (puntkomma, UTF-8 met BOM).

## Copy-ADGroupMembership.ps1

```powershell
.\Copy-ADGroupMembership.ps1                                                                  # vraagt om bron en doel, daarna J/N
.\Copy-ADGroupMembership.ps1 -SourceEmail 'jan@contoso.nl' -TargetEmail 'piet@contoso.nl' -WhatIf # alleen tonen
.\Copy-ADGroupMembership.ps1 -SourceEmail 'jan@contoso.nl' -TargetEmail 'piet@contoso.nl' -Force  # zonder J/N-vraag
```

| Parameter | Omschrijving |
|---|---|
| `-SourceEmail` / `-TargetEmail` | E-mailadressen van bron en doel (worden gevraagd als je ze weglaat). |
| `-WhatIf` | Toont welke groepen toegevoegd zouden worden, zonder iets te wijzigen. De J/N-vraag wordt dan overgeslagen. |
| `-Force` | Slaat de J/N-vraag over. |

## Vereisten
- Module `ActiveDirectory` (RSAT)
- Voor Compare-ADGroupMembership: leesrechten op de groep en gebruikers; voor een Excel-rapport de module `ImportExcel` (`Install-Module ImportExcel -Scope CurrentUser`)
- Voor Copy-ADGroupMembership: rechten om groepsleden te wijzigen
