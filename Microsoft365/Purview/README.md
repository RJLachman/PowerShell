# Microsoft 365 – Purview

Scripts voor Microsoft Purview (Security & Compliance).

| Script | Omschrijving |
|---|---|
| [SoftDelete-EmailFromSender.ps1](SoftDelete-EmailFromSender.ps1) | Zoekt via een Content Search alle mails van een opgegeven afzender in alle mailboxen en kan ze na bevestiging soft-deleten. |

## Vereisten
- Module `ExchangeOnlineManagement` (voor `Connect-IPPSSession`)
- Rollen in Purview: *Compliance Search* en, voor verwijderen, *Search And Purge*
