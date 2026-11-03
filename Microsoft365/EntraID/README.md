# Microsoft 365 – Entra ID (Azure AD)

| Script | Omschrijving |
|---|---|
| [Remove-StaleAzureADDevices.ps1](Remove-StaleAzureADDevices.ps1) | Exporteert alle apparaten die langer dan 60 dagen (instelbaar) niet zijn ingelogd naar een CSV en verwijdert ze na bevestiging. |

## Vereisten
- Module `AzureAD` (let op: deze module is door Microsoft uitgefaseerd; Microsoft Graph PowerShell is de opvolger)
- Rol: Cloud Device Administrator of hoger
