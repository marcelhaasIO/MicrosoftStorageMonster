# Microsoft Storage Monster — PowerShell Scanner

  **Open-source PowerShell script to find and report storage waste in your Microsoft 365 tenant (SharePoint Online & OneDrive for Business) — with zero data leaving your machine.**

  > Part of the [StorageScan](https://storagescan.app) ecosystem. Use this script standalone for free, or optionally upload results to your StorageScan dashboard.

  ---

  ## What It Detects

  | Finding Type | Description |
  |---|---|
  | **Stale Files** | Files not modified in N days (default: 730 / 2 years) |
  | **Large Files** | Individual files exceeding a size threshold (default: 100 MB) |
  | **Version Bloat** | Files where historical versions consume more than N× the current file size |
  | **Duplicates** | Files with identical name and size in the same library |

  ---

  ## Prerequisites

  - PowerShell 5.1+ or PowerShell 7+
  - [PnP.PowerShell](https://pnp.github.io/powershell/) module

  ```powershell
  Install-Module PnP.PowerShell -Scope CurrentUser -Force
  ```

  ---

  ## Quick Start

  **Scan all sites in your tenant (interactive login):**
  ```powershell
  .\StorageScan-Local.ps1
  ```

  **Scan a single site:**
  ```powershell
  .\StorageScan-Local.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/IT"
  ```

  **Custom thresholds + export CSV and JSON:**
  ```powershell
  .\StorageScan-Local.ps1 -StaleThresholdDays 365 -LargeFileMB 50 -ExportCsv -ExportJson -OutputPath "C:\Reports"
  ```

  **App registration (no browser popup — great for automation/CI):**
  ```powershell
  .\StorageScan-Local.ps1 `
      -TenantId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
      -ClientId     "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
      -ClientSecret "your-client-secret" `
      -SiteUrl      "https://contoso.sharepoint.com/sites/IT"
  ```

  **Fully unattended — scan all sites, no browser, no prompts:**
  ```powershell
  .\StorageScan-Local.ps1 `
      -TenantId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
      -ClientId     "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
      -ClientSecret "your-client-secret" `
      -AdminUrl     "https://contoso-admin.sharepoint.com" `
      -ExportCsv -ExportJson -OutputPath "C:\Reports"
  ```

  **Upload results to your StorageScan dashboard:**
  ```powershell
  .\StorageScan-Local.ps1 -ApiKey "ssk_yourkeyhere" -OutputPath "C:\Reports"
  ```

  ---

  ## Parameters

  | Parameter | Type | Default | Description |
  |---|---|---|---|
  | `-SiteUrl` | string | `"all"` | Specific site URL to scan, or `"all"` for the whole tenant |
  | `-StaleThresholdDays` | int | `730` | Days since last modification before a file is flagged as stale |
  | `-LargeFileMB` | int | `100` | File size in MB above which a file is flagged as large |
  | `-VersionBloatMultiplier` | int | `5` | Flag files where version history exceeds current size × this value |
  | `-OutputPath` | string | `.` | Directory where the HTML report will be written |
  | `-ExportCsv` | switch | off | Also write findings to a CSV file |
  | `-ExportJson` | switch | off | Also write findings to a JSON file |
  | `-ApiKey` | string | `""` | StorageScan API key to upload results to your dashboard |
  | `-ApiBaseUrl` | string | `https://storagescan.app/api` | StorageScan API base URL |
  | `-TenantId` | string | `""` | Azure AD Tenant ID for app registration auth |
  | `-ClientId` | string | `""` | Azure AD Application (client) ID for app registration auth |
  | `-ClientSecret` | string | `""` | Client secret for the app registration |
  | `-AdminUrl` | string | `""` | SharePoint Admin Center URL (e.g. `https://contoso-admin.sharepoint.com`). Skips the interactive prompt when scanning all sites |

  ---

  ## Authentication

  ### Interactive (default)

  When you run without `-TenantId`/`-ClientId`/`-ClientSecret`, a browser window opens for Microsoft 365 login. You need:
  - A Microsoft 365 account with **read access** to the sites you want to scan
  - For "scan all sites": **SharePoint Administrator** role to enumerate all site collections

  ### App Registration (recommended for automation)

  Pass `-TenantId`, `-ClientId`, and `-ClientSecret` to authenticate silently — no browser required. Ideal for scheduled tasks, CI/CD pipelines, and server deployments.

  **Required app permissions in Azure AD:**

  | Scenario | Permission required |
  |---|---|
  | Single-site scan | SharePoint → `Sites.Read.All` (Application) |
  | Scan all sites | SharePoint → `Sites.FullControl.All` (Application) |

  **How to set up an app registration:**
  1. Go to [Azure Portal → App registrations](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) → **New registration**
  2. Name it (e.g. `StorageScan-Local`), leave defaults, click **Register**
  3. Note the **Application (client) ID** and **Directory (tenant) ID**
  4. Under **Certificates & secrets** → **New client secret** → copy the value
  5. Under **API permissions** → **Add a permission** → **SharePoint** → **Application permissions**
     - Add `Sites.Read.All` (single site) or `Sites.FullControl.All` (all sites)
  6. Click **Grant admin consent**

  ---

  ## Output

  ### HTML Report
  The script always generates a self-contained HTML report (`StorageScan-Report-YYYYMMDD-HHmmss.html`) with:
  - Summary cards (total findings, potential GB savings, estimated monthly cost savings)
  - Filterable and sortable findings table
  - Filter by finding type, search by filename or path

  ### CSV Export (`-ExportCsv`)
  Flat CSV with all findings including file path, size, last modified date, version counts, and potential savings bytes.

  ### JSON Export (`-ExportJson`)
  Full findings array as JSON, suitable for piping into other tools or dashboards.

  ---

  ## Privacy

  Your data **never leaves your machine** (unless you explicitly use `-ApiKey`).

  - Authentication is handled via Microsoft interactive browser login or your own Azure AD app registration
  - All scanning happens locally via Microsoft Graph / SharePoint REST APIs
  - Your credentials are passed directly to Microsoft — not stored or intercepted
  - The HTML/CSV/JSON outputs stay on your disk

  ---

  ## Optional: Upload to StorageScan Dashboard

  If you have a [StorageScan](https://storagescan.app) account, you can upload your scan results for a richer dashboard view, trend tracking, and AI-powered recommendations:

  ```powershell
  .\StorageScan-Local.ps1 -ApiKey "ssk_yourkeyhere"
  ```

  Get your API key from the StorageScan dashboard under **Settings → API Keys**.

  ---

  ## License

  MIT — see [LICENSE](LICENSE)

  ---

  ## Contributing

  Pull requests welcome. See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.
  