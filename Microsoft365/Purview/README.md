# Microsoft 365 – Purview

Scripts voor Microsoft Purview (Security & Compliance).

| Script | Omschrijving |
|---|---|
| [Remove-EmailFromSender.ps1](Remove-EmailFromSender.ps1) | Zoekt via een Content Search alle mails van een opgegeven afzender in alle mailboxen en kan ze na bevestiging soft-deleten. *(Voorheen `SoftDelete-EmailFromSender.ps1`.)* |

## Gebruik
```powershell
.\Remove-EmailFromSender.ps1                                              # vraagt om afzender, daarna J/N
.\Remove-EmailFromSender.ps1 -SenderAddress 'afzender@voorbeeld.nl' -WhatIf # alleen zoeken, niets verwijderen
.\Remove-EmailFromSender.ps1 -SenderAddress 'afzender@voorbeeld.nl' -Force  # zonder J/N-vraag verwijderen
```

| Parameter | Omschrijving |
|---|---|
| `-SenderAddress` | E-mailadres van de afzender (wordt gevraagd als je het weglaat). |
| `-PollIntervalSeconds` | Seconden tussen statuscontroles (standaard 10). |
| `-WhatIf` | Voert de search uit en toont wat er verwijderd zou worden, zonder te verwijderen. De J/N-vraag wordt dan overgeslagen. |
| `-Force` | Slaat de J/N-vraag over. |

> Soft-deleted items komen in *Herstelbare items* en zijn binnen de bewaartermijn te herstellen. Een purge verwijdert maximaal 10 items per mailbox per actie (limiet van Microsoft).

## Vereisten
- Module `ExchangeOnlineManagement` (voor `Connect-IPPSSession`)
- Rollen in Purview: *Compliance Search* en, voor verwijderen, *Search And Purge*
