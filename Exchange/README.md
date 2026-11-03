# Exchange (hybrid / on-premises)

Scripts voor Exchange on-premises in een hybride omgeving met Exchange Online.

| Script | Omschrijving |
|---|---|
| [New-SharedRemoteMailbox.ps1](New-SharedRemoteMailbox.ps1) | Maakt een shared remote mailbox aan, genereert alias en UPN, controleert of het adres al bestaat en zet de AD Description op "Shared Mailbox - <naam>". |

## Gebruik
```powershell
.\New-SharedRemoteMailbox.ps1 -Domain "jouwdomein.nl" -OU "jouwdomein.local/Shared Mailboxes"
```
De standaardwaarden `contoso.nl` en `contoso.local/Shared Mailboxes` zijn voorbeelden; geef je eigen domein en OU mee of pas de defaults aan.

## Vereisten
- Exchange Management Shell (on-premises)
- Optioneel: module `ActiveDirectory` (RSAT); anders wordt ADSI gebruikt
