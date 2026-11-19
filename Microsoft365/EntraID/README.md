# Microsoft 365 – Entra ID (Azure AD)

| Script | Omschrijving |
|---|---|
| [Remove-StaleAzureADDevices.ps1](Remove-StaleAzureADDevices.ps1) | Exporteert alle apparaten die langer dan 60 dagen (instelbaar) niet zijn aangemeld naar een CSV en verwijdert ze na bevestiging. Gebruikt Microsoft Graph. |

## Gebruik
```powershell
.\Remove-StaleAzureADDevices.ps1                                  # 60 dagen, export + J/N
.\Remove-StaleAzureADDevices.ps1 -DeletionThresholdDays 90 -WhatIf # toont wat er verwijderd zou worden
.\Remove-StaleAzureADDevices.ps1 -ExportPath 'C:\temp\stale.csv' -Force
```

| Parameter | Omschrijving |
|---|---|
| `-DeletionThresholdDays` | Dagen zonder aanmelding (standaard 60). Apparaten die nooit zijn aangemeld worden overgeslagen. |
| `-ExportPath` | Pad van de CSV-export (standaard `EntraDeviceExport.csv` naast het script). |
| `-WhatIf` | Maakt wel de export, maar verwijdert niets. De J/N-vraag wordt dan overgeslagen. |
| `-Force` | Slaat de J/N-vraag over. |

De login gebruikt geen WAM (voorkomt de fout *"A window handle must be configured"*) en valt zo nodig terug op device code login.

## Vereisten
- `Install-Module Microsoft.Graph.Identity.DirectoryManagement`
- Scope `Device.ReadWrite.All` (rol Cloud Device Administrator of hoger)
