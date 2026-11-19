# Exchange (hybrid / on-premises)

Scripts voor Exchange on-premises in een hybride omgeving met Exchange Online.

| Script | Omschrijving |
|---|---|
| [New-SharedRemoteMailbox.ps1](New-SharedRemoteMailbox.ps1) | Maakt een shared remote mailbox aan, genereert alias en UPN, controleert of het adres al bestaat en zet de AD Description op "Shared Mailbox - <naam>". |

## Gebruik
```powershell
.\New-SharedRemoteMailbox.ps1 -Domain 'jouwdomein.nl' -OU 'jouwdomein.local/Shared Mailboxes'
.\New-SharedRemoteMailbox.ps1 -DisplayName 'Afdeling Inkoop' -WhatIf              # alleen tonen wat er zou gebeuren
.\New-SharedRemoteMailbox.ps1 -DisplayName 'Afdeling Inkoop' -Alias 'inkoop' -Force # zonder J/N-vraag
```
De standaardwaarden `contoso.nl` en `contoso.local/Shared Mailboxes` zijn voorbeelden; geef je eigen domein en OU mee of pas de defaults aan.

| Parameter | Omschrijving |
|---|---|
| `-Domain` | E-maildomein voor UPN en primair SMTP-adres (standaard `contoso.nl`). |
| `-OU` | On-prem OU voor het account (standaard `contoso.local/Shared Mailboxes`). |
| `-DisplayName` / `-Alias` | Weergavenaam en alias (worden gevraagd als je ze weglaat). |
| `-WhatIf` | Toont wat er aangemaakt zou worden, zonder iets te wijzigen. De J/N-vraag wordt dan overgeslagen. |
| `-Force` | Slaat de J/N-vraag over. |

## Vereisten
- Exchange Management Shell (on-premises)
- Optioneel: module `ActiveDirectory` (RSAT); anders wordt ADSI gebruikt
