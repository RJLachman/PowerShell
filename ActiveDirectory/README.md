# Active Directory (on-premises)

Scripts voor lokale Active Directory.

| Script | Omschrijving |
|---|---|
| [Copy-ADGroupMembership.ps1](Copy-ADGroupMembership.ps1) | Kopieert de groepslidmaatschappen van een brongebruiker naar een doelgebruiker, op basis van e-mailadres. Toont eerst welke groepen erbij komen en vraagt om bevestiging. |

## Gebruik
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
- Rechten om groepsleden te wijzigen
