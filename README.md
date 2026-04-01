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

**Scan all sites in your tenant:**
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

- Authentication is handled via Microsoft's interactive browser login (PnP.PowerShell)
- All scanning happens locally via Microsoft Graph / SharePoint REST APIs
- Your credentials are passed directly to Microsoft — not stored or intercepted
- The HTML/CSV/JSON outputs stay on your disk

---

## Authentication

When you run the script, a browser window will open for Microsoft 365 interactive login. You need:
- A Microsoft 365 account with **read access** to the sites you want to scan
- For "scan all sites": **SharePoint Administrator** role to enumerate all site collections

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
