# Microsoft 365 – Entra ID (Azure AD)

| Script | Omschrijving |
|---|---|
| [Remove-StaleAzureADDevices.ps1](Remove-StaleAzureADDevices.ps1) | Exporteert alle apparaten die langer dan 60 dagen (instelbaar met `-DeletionThresholdDays`) niet zijn ingelogd naar een CSV en verwijdert ze na bevestiging. |

## Vereisten
- `Install-Module Microsoft.Graph.Identity.DirectoryManagement`
- Scope `Device.ReadWrite.All` (rol Cloud Device Administrator of hoger)
