<#
.SYNOPSIS
    StorageScan Local - SharePoint Storage Waste Scanner (Privacy-First, No Cloud Required)

.DESCRIPTION
    Scans your Microsoft 365 SharePoint Online and OneDrive for Business tenant for storage
    waste: version bloat, stale files, large files, and duplicate candidates.

    Runs entirely on your own machine using your own credentials via Microsoft Graph.
    No data is sent to any third-party service unless you explicitly use -ApiKey to upload
    results to your StorageScan dashboard.

    Requires: PowerShell 5.1+ or PowerShell 7+, PnP.PowerShell module.

.PARAMETER SiteUrl
    The URL of a specific SharePoint site to scan (e.g. https://contoso.sharepoint.com/sites/IT).
    Use "all" or omit to scan all sites in the tenant.

.PARAMETER StaleThresholdDays
    Number of days since last modification after which a file is considered stale.
    Default: 730 (2 years).

.PARAMETER LargeFileMB
    Size in megabytes above which a file is flagged as a large file.
    Default: 100 MB.

.PARAMETER VersionBloatMultiplier
    Flag files where total version size exceeds (current size x multiplier).
    Default: 5.

.PARAMETER OutputPath
    Directory where the HTML report (and optional CSV/JSON) will be written.
    Default: current directory.

.PARAMETER ExportCsv
    If specified, also writes findings to a CSV file.

.PARAMETER ExportJson
    If specified, also writes findings to a JSON file.

.PARAMETER ApiKey
    Optional StorageScan API key. If provided, results are uploaded to your StorageScan
    dashboard after the local scan completes.

.PARAMETER ApiBaseUrl
    Base URL of the StorageScan API. Default: https://storagescan.app/api

.PARAMETER TenantId
    Azure AD Tenant ID for app registration authentication.
    Must be combined with -ClientId and -ClientSecret.
    When all three are provided, interactive browser login is skipped entirely.

.PARAMETER ClientId
    Azure AD Application (client) ID for app registration authentication.

.PARAMETER ClientSecret
    Client secret for the Azure AD app registration.
    The app registration requires the SharePoint "Sites.Read.All" app permission
    (and "Sites.FullControl.All" if scanning all sites via the admin center).

.PARAMETER AdminUrl
    SharePoint Admin Center URL (e.g. https://contoso-admin.sharepoint.com).
    Only used when -SiteUrl is "all". If omitted, the script will prompt for it.
    Providing this enables fully headless / unattended runs when combined with app reg auth.

.EXAMPLE
    .\StorageScan-Local.ps1

.EXAMPLE
    .\StorageScan-Local.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/IT" -LargeFileMB 50

.EXAMPLE
    .\StorageScan-Local.ps1 -StaleThresholdDays 365 -ExportCsv -ExportJson -OutputPath "C:\Reports"

.EXAMPLE
    .\StorageScan-Local.ps1 -ApiKey "ssk_yourkeyhere" -OutputPath "C:\Reports"

.EXAMPLE
    .\StorageScan-Local.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
                             -ClientId  "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
                             -ClientSecret "your-client-secret" `
                             -OutputPath "C:\Reports"

.EXAMPLE
    .\StorageScan-Local.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
                             -ClientId  "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
                             -ClientSecret "your-client-secret" `
                             -SiteUrl "https://contoso.sharepoint.com/sites/IT" `
                             -ExportCsv

.EXAMPLE
    # Fully unattended — scan all sites with no browser and no prompts
    .\StorageScan-Local.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
                             -ClientId  "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
                             -ClientSecret "your-client-secret" `
                             -AdminUrl "https://contoso-admin.sharepoint.com" `
                             -SiteUrl "all" `
                             -ExportCsv -ExportJson

.NOTES
    Install PnP.PowerShell before running:
        Install-Module PnP.PowerShell -Scope CurrentUser -Force

    Authentication modes:
      Interactive (default) — opens a browser window for Microsoft 365 login.
      App registration       — pass -TenantId, -ClientId, and -ClientSecret to skip
                               the browser entirely (ideal for automation / CI).

    Required app permissions for app registration mode:
      SharePoint > Sites.Read.All          (for single-site scans)
      SharePoint > Sites.FullControl.All   (for scanning all sites via admin center)
#>

[CmdletBinding()]
param(
    [string]$SiteUrl = "all",
    [int]$StaleThresholdDays = 730,
    [int]$LargeFileMB = 100,
    [int]$VersionBloatMultiplier = 5,
    [string]$OutputPath = ".",
    [switch]$ExportCsv,
    [switch]$ExportJson,
    [string]$ApiKey = "",
    [string]$ApiBaseUrl = "https://storagescan.app/api",

    # App registration authentication (alternative to interactive browser login)
    [string]$TenantId = "",
    [string]$ClientId = "",
    [string]$ClientSecret = "",

    # SharePoint Admin Center URL — required when SiteUrl is "all".
    # If omitted, the script will prompt for it interactively.
    # E.g. https://contoso-admin.sharepoint.com
    [string]$AdminUrl = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Helpers

function Write-Header {
    $orange = "Red"
    $gold   = "Yellow"
    $blue   = "Cyan"
    $dim    = "DarkGray"

    Clear-Host

    # ── Monster (single-quoted here-string = no escaping needed) ────────────
    $monsterArt = @'

                   .~~~~~~~~~~~~~~~~~~~~~~~~~.
                  /  (@) (@) (@) (@) (@) (@)   \
                 |  (@) (@) (@) (@) (@) (@) (@)  |
                 |       .~~~~~~~~~~~~~~~~.      |
                 |      | W W W W W W W W W |    |
                 |       '~~~~~~~~~~~~~~~~'      |
                  \           ~ CHOMP ~          /
                   '~~~~~~~~~~~~~~~~~~~~~~~~~'
                       |    |    |    |    |
                      /|    |    |    |    |\
                     /_|____|____|____|____|_\

'@
    Write-Host $monsterArt -ForegroundColor $orange

    # ── ASCII-art title (figlet "standard" font) ─────────────────────────────
    $titleArt = @'
   ____  _                                 ____
  / ___|| |_ ___  _ __ __ _  __ _  ___   / ___|  ___ __ _ _ __
  \___ \| __/ _ \| '__/ _` |/ _` |/ _ \  \___ \ / __/ _` | '_ \
   ___) | || (_) | | | (_| | (_| |  __/   ___) | (_| (_| | | | |
  |____/ \__\___/|_|  \__,_|\__, |\___|  |____/ \___\__,_|_| |_|
                             |___/

'@
    Write-Host $titleArt -ForegroundColor $gold

    # ── Subtitle bar ─────────────────────────────────────────────────────────
    $sep = '-' * 67
    Write-Host "  $sep" -ForegroundColor $dim
    Write-Host '   L O C A L  SCANNER   |  Microsoft 365 Storage Waste Analyzer  ' -ForegroundColor $blue
    Write-Host '   Privacy-First  *  No cloud required  *  PnP.PowerShell powered ' -ForegroundColor $dim
    Write-Host "  $sep" -ForegroundColor $dim
    Write-Host ""
}

function Write-Step {
    param([string]$Message)
    Write-Host "[*] $Message" -ForegroundColor Yellow
}

function Write-Success {
    param([string]$Message)
    Write-Host "[+] $Message" -ForegroundColor Green
}

function Write-Info {
    param([string]$Message)
    Write-Host "    $Message" -ForegroundColor Gray
}

function Format-Bytes {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes B"
}

function Ensure-PnPModule {
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        Write-Host "[!] PnP.PowerShell module not found." -ForegroundColor Red
        Write-Host "    Install it with: Install-Module PnP.PowerShell -Scope CurrentUser -Force" -ForegroundColor Red
        exit 1
    }
    Import-Module PnP.PowerShell -ErrorAction Stop
}

function Connect-ToSite {
    <#
    .SYNOPSIS
        Connect to a SharePoint site using either app registration or interactive auth.
    #>
    param(
        [Parameter(Mandatory)][string]$Url,
        [switch]$SilentlyContinue
    )

    $ea = if ($SilentlyContinue) { "SilentlyContinue" } else { "Stop" }

    if ($script:ClientId -ne "" -and $script:ClientSecret -ne "" -and $script:TenantId -ne "") {
        Connect-PnPOnline -Url $Url `
            -ClientId     $script:ClientId `
            -ClientSecret $script:ClientSecret `
            -Tenant       $script:TenantId `
            -ErrorAction  $ea
    } else {
        Connect-PnPOnline -Url $Url -Interactive -ErrorAction $ea
    }
}

function Test-AppRegAuth {
    return ($script:ClientId -ne "" -and $script:ClientSecret -ne "" -and $script:TenantId -ne "")
}

#endregion

#region Scanning

function Get-AllSiteUrls {
    param([string]$AdminUrl)
    Write-Step "Enumerating all SharePoint site collections..."
    $sites = Get-PnPTenantSite -Connection (Get-PnPConnection) -ErrorAction Stop
    return $sites | ForEach-Object { $_.Url }
}

function Get-SiteDriveItems {
    param(
        [string]$SiteWebUrl,
        [int]$StaleThresholdDays,
        [long]$LargeFileBytes,
        [int]$VersionBloatMultiplier
    )

    $findings = [System.Collections.Generic.List[hashtable]]::new()
    $now = Get-Date

    try {
        Connect-ToSite -Url $SiteWebUrl
    } catch {
        Write-Info "  Could not connect to $SiteWebUrl - skipping. ($_)"
        return $findings
    }

    try {
        $lists = Get-PnPList -ErrorAction Stop | Where-Object { $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $false }
    } catch {
        Write-Info "  Could not list libraries on $SiteWebUrl - skipping."
        return $findings
    }

    foreach ($list in $lists) {
        Write-Info "  Scanning library: $($list.Title)"

        try {
            $items = Get-PnPListItem -List $list -Fields "FileLeafRef","FileRef","File_x0020_Size","Modified","FSObjType" -PageSize 500 -ErrorAction Stop |
                     Where-Object { $_["FSObjType"] -eq 0 }
        } catch {
            Write-Info "    Could not enumerate items in $($list.Title) - skipping."
            continue
        }

        $hashGroups = @{}

        foreach ($item in $items) {
            $fileName  = $item["FileLeafRef"]
            $filePath  = $item["FileRef"]
            $fileSize  = [long]($item["File_x0020_Size"] ?? 0)
            $modified  = $item["Modified"]

            $daysSinceMod = if ($modified) { ($now - $modified).Days } else { 0 }

            if ($modified -and $daysSinceMod -gt $StaleThresholdDays) {
                $findings.Add(@{
                    FindingType         = "stale"
                    SiteUrl             = $SiteWebUrl
                    LibraryName         = $list.Title
                    FileName            = $fileName
                    FilePath            = $filePath
                    FileSizeBytes       = $fileSize
                    LastModified        = if ($modified) { $modified.ToString("o") } else { "" }
                    DaysSinceModified   = $daysSinceMod
                    VersionsCount       = 0
                    VersionTotalSizeBytes = 0
                    PotentialSavingsBytes = $fileSize
                    DuplicateGroup      = ""
                    Details             = "Not modified in $daysSinceMod days"
                })
            }

            if ($fileSize -gt $LargeFileBytes) {
                $findings.Add(@{
                    FindingType         = "large_file"
                    SiteUrl             = $SiteWebUrl
                    LibraryName         = $list.Title
                    FileName            = $fileName
                    FilePath            = $filePath
                    FileSizeBytes       = $fileSize
                    LastModified        = if ($modified) { $modified.ToString("o") } else { "" }
                    DaysSinceModified   = $daysSinceMod
                    VersionsCount       = 0
                    VersionTotalSizeBytes = 0
                    PotentialSavingsBytes = $fileSize
                    DuplicateGroup      = ""
                    Details             = "File size: $(Format-Bytes $fileSize)"
                })
            }

            $dupKey = "$fileSize-$fileName"
            if (-not $hashGroups.ContainsKey($dupKey)) { $hashGroups[$dupKey] = [System.Collections.Generic.List[string]]::new() }
            $hashGroups[$dupKey].Add($filePath)

            try {
                $versions = Get-PnPProperty -ClientObject $item.File -Property Versions -ErrorAction Stop
                $versionCount = $versions.Count
                if ($versionCount -gt 1) {
                    $versionTotalSize = ($versions | ForEach-Object { [long]($_.Size ?? 0) } | Measure-Object -Sum).Sum
                    if ($versionTotalSize -gt ($fileSize * $VersionBloatMultiplier)) {
                        $savings = $versionTotalSize - $fileSize
                        $findings.Add(@{
                            FindingType         = "version_bloat"
                            SiteUrl             = $SiteWebUrl
                            LibraryName         = $list.Title
                            FileName            = $fileName
                            FilePath            = $filePath
                            FileSizeBytes       = $fileSize
                            LastModified        = if ($modified) { $modified.ToString("o") } else { "" }
                            DaysSinceModified   = $daysSinceMod
                            VersionsCount       = $versionCount
                            VersionTotalSizeBytes = $versionTotalSize
                            PotentialSavingsBytes = $savings
                            DuplicateGroup      = ""
                            Details             = "$versionCount versions, total $(Format-Bytes $versionTotalSize)"
                        })
                    }
                }
            } catch {
                # Version history not accessible for this item; skip silently
            }
        }

        foreach ($entry in $hashGroups.GetEnumerator()) {
            if ($entry.Value.Count -gt 1) {
                $dupGroup = $entry.Key
                foreach ($path in $entry.Value) {
                    $matchingItem = $items | Where-Object { $_["FileRef"] -eq $path } | Select-Object -First 1
                    $fSize = if ($matchingItem) { [long]($matchingItem["File_x0020_Size"] ?? 0) } else { 0 }
                    $fName = if ($matchingItem) { $matchingItem["FileLeafRef"] } else { Split-Path $path -Leaf }
                    $fMod  = if ($matchingItem) { $matchingItem["Modified"] } else { $null }
                    $findings.Add(@{
                        FindingType         = "duplicate"
                        SiteUrl             = $SiteWebUrl
                        LibraryName         = $list.Title
                        FileName            = $fName
                        FilePath            = $path
                        FileSizeBytes       = $fSize
                        LastModified        = if ($fMod) { $fMod.ToString("o") } else { "" }
                        DaysSinceModified   = 0
                        VersionsCount       = 0
                        VersionTotalSizeBytes = 0
                        PotentialSavingsBytes = $fSize
                        DuplicateGroup      = $dupGroup
                        Details             = "Duplicate group: $dupGroup ($($entry.Value.Count) copies)"
                    })
                }
            }
        }
    }

    return $findings
}

#endregion

#region HTML Report Generation

function New-HtmlReport {
    param(
        [System.Collections.Generic.List[hashtable]]$Findings,
        [hashtable]$Summary,
        [string]$OutputPath,
        [string]$ScannedScope
    )

    $now = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $totalBytes = $Summary.TotalSavingsBytes
    $gbSavings  = [math]::Round($totalBytes / 1GB, 2)
    $costSavings = [math]::Round(($totalBytes / 1GB) * 0.20, 2)

    $countByType = @{ stale = 0; large_file = 0; version_bloat = 0; duplicate = 0 }
    $savingsByType = @{ stale = 0L; large_file = 0L; version_bloat = 0L; duplicate = 0L }
    foreach ($f in $Findings) {
        $t = $f.FindingType
        if ($countByType.ContainsKey($t)) { $countByType[$t]++ }
        if ($savingsByType.ContainsKey($t)) { $savingsByType[$t] += $f.PotentialSavingsBytes }
    }

    $rowsHtml = ($Findings | ForEach-Object {
        $badgeClass = switch ($_.FindingType) {
            "stale"         { "badge badge-stale" }
            "large_file"    { "badge badge-large" }
            "version_bloat" { "badge badge-version" }
            "duplicate"     { "badge badge-duplicate" }
            default         { "badge" }
        }
        $label = switch ($_.FindingType) {
            "stale"        { "Stale" }
            "large_file"   { "Large File" }
            "version_bloat"{ "Version Bloat" }
            "duplicate"    { "Duplicate" }
            default        { $_.FindingType }
        }
        $savings = Format-Bytes $_.PotentialSavingsBytes
        $size    = Format-Bytes $_.FileSizeBytes
        $modDate = if ($_.LastModified) { $_.LastModified.Substring(0,10) } else { "—" }
        $path    = $_.SiteUrl -replace 'https://[^/]+', ''

        "<tr>
            <td><span class='$badgeClass' data-type='$($_.FindingType)'>$label</span></td>
            <td class='fname' title='$($_.FilePath)'>$($_.FileName)</td>
            <td class='muted sm' title='$($_.FilePath)'>$path</td>
            <td>$size</td>
            <td class='muted'>$modDate</td>
            <td class='muted'>$($_.VersionsCount)</td>
            <td class='savings'>$savings</td>
            <td class='muted sm'>$($_.Details)</td>
        </tr>"
    }) -join "`n"

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1.0"/>
  <title>StorageScan Local Report &mdash; $now</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    :root{
      --bg:#0d0d12;--surface:#13131a;--surface2:#1a1a24;--border:#21212e;
      --text:#e2e2ea;--muted:#6b6b7e;
      --orange:#CF4B00;--orange-bg:#1c0e06;--orange-border:#2e1a0c;
      --blue:#9CC6DB;--blue-bg:#0a1620;--blue-border:#0f2030;
      --gold:#DDBA7D;--gold-bg:#201808;--gold-border:#32280e;
      --green:#22c55e;--red:#ef4444;--red-bg:#1a0808;--red-border:#2e1010;
    }
    body{font-family:'Inter',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;-webkit-font-smoothing:antialiased}
    /* ── Top bar ── */
    .topbar{height:3px;background:linear-gradient(90deg,var(--orange) 0%,var(--gold) 50%,var(--blue) 100%)}
    /* ── Header ── */
    .header{background:var(--surface);border-bottom:1px solid var(--border);padding:26px 40px;display:flex;align-items:center;gap:18px}
    .header-icon{font-size:36px;line-height:1;flex-shrink:0;filter:grayscale(0.1)}
    .header-text h1{font-size:21px;font-weight:800;color:#fff;letter-spacing:-0.3px}
    .header-text h1 .brand{color:var(--orange)}
    .header-text p{color:var(--muted);margin-top:5px;font-size:12.5px;line-height:1.6}
    /* ── Layout ── */
    .container{max-width:1440px;margin:0 auto;padding:32px 40px}
    /* ── Summary cards ── */
    .cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(185px,1fr));gap:14px;margin-bottom:32px}
    .card{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:18px 20px;position:relative;overflow:hidden}
    .card::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;border-radius:12px 12px 0 0;background:var(--accent,var(--border))}
    .c-total{--accent:var(--orange)}
    .c-save {--accent:var(--green)}
    .c-stale{--accent:var(--gold)}
    .c-large{--accent:var(--blue)}
    .c-ver  {--accent:var(--orange)}
    .c-dup  {--accent:var(--red)}
    .card .lbl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:var(--muted);margin-bottom:10px}
    .card .val{font-size:30px;font-weight:800;line-height:1;color:#fff}
    .val.c-orange{color:var(--orange)} .val.c-green{color:var(--green)}
    .val.c-gold{color:var(--gold)}     .val.c-blue{color:var(--blue)}
    .card .sub{font-size:11.5px;color:var(--muted);margin-top:6px}
    /* ── Section ── */
    .section{font-size:15px;font-weight:700;color:#fff;margin-bottom:14px;display:flex;align-items:center;gap:8px}
    .section::before{content:'';display:inline-block;width:3px;height:15px;background:var(--orange);border-radius:2px;flex-shrink:0}
    /* ── Filters ── */
    .filters{display:flex;gap:10px;margin-bottom:12px;flex-wrap:wrap;align-items:center}
    .filters input,.filters select{
      background:var(--surface);border:1px solid var(--border);border-radius:8px;
      color:var(--text);padding:7px 12px;font-size:13px;font-family:inherit;outline:none
    }
    .filters input{flex:1;min-width:220px}
    .filters input::placeholder{color:var(--muted)}
    .filters input:focus,.filters select:focus{border-color:var(--orange);box-shadow:0 0 0 3px var(--orange-bg)}
    .filters select option{background:var(--surface)}
    /* ── Pills ── */
    .pills{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:16px}
    .pill{padding:4px 13px;border-radius:9999px;font-size:12px;font-weight:600;cursor:pointer;
      border:1px solid var(--border);background:var(--surface);color:var(--muted);
      transition:all .15s;user-select:none}
    .pill:hover,.pill.active{border-color:var(--orange);color:var(--orange);background:var(--orange-bg)}
    /* ── Table ── */
    .tbl-wrap{overflow-x:auto;border-radius:12px;border:1px solid var(--border);margin-bottom:32px}
    table{width:100%;border-collapse:collapse;font-size:13px}
    thead{background:var(--surface)}
    thead th{padding:11px 14px;text-align:left;font-size:10px;font-weight:700;text-transform:uppercase;
      letter-spacing:.08em;color:var(--muted);border-bottom:1px solid var(--border);
      cursor:pointer;user-select:none;white-space:nowrap}
    thead th:hover{color:var(--orange)}
    thead th::after{content:' \25B4\25BE';opacity:.25;font-size:8px}
    tbody tr{border-bottom:1px solid #15151e;transition:background .1s}
    tbody tr:hover{background:var(--surface2)}
    tbody td{padding:10px 14px;color:#c4c4d2;vertical-align:middle;max-width:260px;
      overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    td.muted{color:var(--muted)}
    td.sm{font-size:12px}
    td.fname{color:var(--text);font-weight:500}
    td.savings{color:var(--green);font-weight:600}
    /* ── Badges ── */
    .badge{display:inline-flex;align-items:center;padding:2px 9px;border-radius:9999px;font-size:11px;font-weight:600;white-space:nowrap;border:1px solid transparent}
    .badge-stale  {background:var(--gold-bg);color:var(--gold);border-color:var(--gold-border)}
    .badge-large  {background:var(--blue-bg);color:var(--blue);border-color:var(--blue-border)}
    .badge-version{background:var(--orange-bg);color:var(--orange);border-color:var(--orange-border)}
    .badge-duplicate{background:var(--red-bg);color:var(--red);border-color:var(--red-border)}
    /* ── Footer ── */
    .footer{text-align:center;color:var(--muted);font-size:12px;padding:28px 40px;border-top:1px solid var(--border)}
    .footer a{color:var(--orange);text-decoration:none}
    .footer a:hover{text-decoration:underline}
    .no-results{text-align:center;padding:52px;color:var(--muted);font-size:14px}
    @media(max-width:768px){
      .container,.header{padding:20px 16px}
      .card .val{font-size:24px}
      .header{flex-direction:column;align-items:flex-start}
    }
  </style>
</head>
<body>
<div class="topbar"></div>
<div class="header">
  <div class="header-icon">&#129681;</div>
  <div class="header-text">
    <h1><span class="brand">StorageScan</span> &mdash; Local Storage Waste Report</h1>
    <p>
      Generated: $now &nbsp;&bull;&nbsp; Scope: $ScannedScope &nbsp;&bull;&nbsp;
      Stale &gt; $StaleThresholdDays days &nbsp;&bull;&nbsp;
      Large &gt; $LargeFileMB MB &nbsp;&bull;&nbsp;
      Version Bloat &gt; ${VersionBloatMultiplier}x
    </p>
  </div>
</div>
<div class="container">
  <div class="cards">
    <div class="card c-total">
      <div class="lbl">Total Findings</div>
      <div class="val c-orange">$($Findings.Count)</div>
      <div class="sub">Across all categories</div>
    </div>
    <div class="card c-save">
      <div class="lbl">Potential Savings</div>
      <div class="val c-green">$gbSavings GB</div>
      <div class="sub">&asymp; `$$costSavings / mo @ `$0.20/GB</div>
    </div>
    <div class="card c-stale">
      <div class="lbl">Stale Files</div>
      <div class="val c-gold">$($countByType.stale)</div>
      <div class="sub">$(Format-Bytes $savingsByType.stale) recoverable</div>
    </div>
    <div class="card c-large">
      <div class="lbl">Large Files</div>
      <div class="val c-blue">$($countByType.large_file)</div>
      <div class="sub">$(Format-Bytes $savingsByType.large_file) total</div>
    </div>
    <div class="card c-ver">
      <div class="lbl">Version Bloat</div>
      <div class="val c-orange">$($countByType.version_bloat)</div>
      <div class="sub">$(Format-Bytes $savingsByType.version_bloat) recoverable</div>
    </div>
    <div class="card c-dup">
      <div class="lbl">Duplicates</div>
      <div class="val">$($countByType.duplicate)</div>
      <div class="sub">$(Format-Bytes $savingsByType.duplicate) recoverable</div>
    </div>
  </div>

  <div class="section">Findings</div>
  <div class="filters">
    <input type="text" id="searchInput" placeholder="Search by file name or path..." oninput="filterTable()">
    <select id="typeFilter" onchange="syncPills()">
      <option value="">All Types</option>
      <option value="stale">Stale</option>
      <option value="large_file">Large File</option>
      <option value="version_bloat">Version Bloat</option>
      <option value="duplicate">Duplicate</option>
    </select>
  </div>
  <div class="pills">
    <span class="pill active" onclick="setPill(this,'')">All</span>
    <span class="pill" onclick="setPill(this,'stale')">Stale</span>
    <span class="pill" onclick="setPill(this,'large_file')">Large Files</span>
    <span class="pill" onclick="setPill(this,'version_bloat')">Version Bloat</span>
    <span class="pill" onclick="setPill(this,'duplicate')">Duplicates</span>
  </div>
  <div class="tbl-wrap">
    <table id="findingsTable">
      <thead>
        <tr>
          <th onclick="sortTable(0)">Type</th>
          <th onclick="sortTable(1)">File Name</th>
          <th onclick="sortTable(2)">Path</th>
          <th onclick="sortTable(3)">Size</th>
          <th onclick="sortTable(4)">Last Modified</th>
          <th onclick="sortTable(5)">Versions</th>
          <th onclick="sortTable(6)">Potential Savings</th>
          <th onclick="sortTable(7)">Details</th>
        </tr>
      </thead>
      <tbody id="tableBody">
$rowsHtml
      </tbody>
    </table>
    <div class="no-results" id="noResults" style="display:none">No findings match your filter.</div>
  </div>
</div>
<div class="footer">
  Generated by <a href="https://storagescan.app" target="_blank">StorageScan Local</a>
  &nbsp;&bull;&nbsp; Your data never left your machine
  &nbsp;&bull;&nbsp; <a href="https://github.com/marcelhaasIO/MicrosoftStorageMonster" target="_blank">Open Source on GitHub</a>
</div>
<script>
  let sortDir = {};
  let activeType = '';

  function filterTable() {
    const q    = document.getElementById('searchInput').value.toLowerCase();
    const type = activeType;
    const rows = document.querySelectorAll('#tableBody tr');
    let vis = 0;
    rows.forEach(r => {
      const dt  = r.querySelector('[data-type]')?.dataset.type || '';
      const txt = r.innerText.toLowerCase();
      const ok  = (!q || txt.includes(q)) && (!type || dt === type);
      r.style.display = ok ? '' : 'none';
      if (ok) vis++;
    });
    document.getElementById('noResults').style.display = vis === 0 ? 'block' : 'none';
  }

  function setPill(el, type) {
    document.querySelectorAll('.pill').forEach(p => p.classList.remove('active'));
    el.classList.add('active');
    activeType = type;
    document.getElementById('typeFilter').value = type;
    filterTable();
  }

  function syncPills() {
    const val = document.getElementById('typeFilter').value;
    document.querySelectorAll('.pill').forEach(p => {
      p.classList.toggle('active', p.onclick.toString().includes("'" + val + "'") || (val === '' && p.onclick.toString().includes("''")));
    });
    activeType = val;
    filterTable();
  }

  function sortTable(col) {
    const tbody = document.getElementById('tableBody');
    const rows  = Array.from(tbody.querySelectorAll('tr'));
    const asc   = !sortDir[col];
    sortDir = {};
    sortDir[col] = asc;
    rows.sort((a, b) => {
      const av = a.cells[col]?.innerText.trim() || '';
      const bv = b.cells[col]?.innerText.trim() || '';
      const an = parseFloat(av.replace(/[^0-9.]/g,''));
      const bn = parseFloat(bv.replace(/[^0-9.]/g,''));
      if (!isNaN(an) && !isNaN(bn)) return asc ? an - bn : bn - an;
      return asc ? av.localeCompare(bv) : bv.localeCompare(av);
    });
    rows.forEach(r => tbody.appendChild(r));
  }
</script>
</body>
</html>
"@

    $reportFile = Join-Path $OutputPath "StorageScan-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
    $html | Out-File -FilePath $reportFile -Encoding UTF8
    return $reportFile
}

#endregion

#region CSV / JSON Export

function Export-FindingsCsv {
    param([System.Collections.Generic.List[hashtable]]$Findings, [string]$OutputPath)
    $csvFile = Join-Path $OutputPath "StorageScan-Findings-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
    $Findings | ForEach-Object {
        [PSCustomObject]@{
            FindingType           = $_.FindingType
            SiteUrl               = $_.SiteUrl
            LibraryName           = $_.LibraryName
            FileName              = $_.FileName
            FilePath              = $_.FilePath
            FileSizeBytes         = $_.FileSizeBytes
            LastModified          = $_.LastModified
            DaysSinceModified     = $_.DaysSinceModified
            VersionsCount         = $_.VersionsCount
            VersionTotalSizeBytes = $_.VersionTotalSizeBytes
            PotentialSavingsBytes = $_.PotentialSavingsBytes
            DuplicateGroup        = $_.DuplicateGroup
            Details               = $_.Details
        }
    } | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
    return $csvFile
}

function Export-FindingsJson {
    param([System.Collections.Generic.List[hashtable]]$Findings, [string]$OutputPath)
    $jsonFile = Join-Path $OutputPath "StorageScan-Findings-$(Get-Date -Format 'yyyyMMdd-HHmmss').json"
    $Findings | ConvertTo-Json -Depth 5 | Out-File -FilePath $jsonFile -Encoding UTF8
    return $jsonFile
}

#endregion

#region Upload to StorageScan

function Upload-ToStorageScan {
    param(
        [System.Collections.Generic.List[hashtable]]$Findings,
        [hashtable]$Summary,
        [string]$ApiKey,
        [string]$ApiBaseUrl,
        [string]$ScannedScope
    )

    Write-Step "Uploading results to StorageScan dashboard..."

    $payload = @{
        source     = "local_powershell"
        scannedAt  = (Get-Date -Format "o")
        scope      = $ScannedScope
        summary    = $Summary
        findings   = ($Findings | ForEach-Object { $_ })
    } | ConvertTo-Json -Depth 10

    try {
        $response = Invoke-RestMethod `
            -Method Post `
            -Uri "$ApiBaseUrl/scan/upload-local" `
            -Headers @{ "Authorization" = "Bearer $ApiKey"; "Content-Type" = "application/json" } `
            -Body $payload `
            -ErrorAction Stop

        Write-Success "Results uploaded. Scan ID: $($response.scanId)"
    } catch {
        Write-Host "[!] Upload failed: $_" -ForegroundColor Red
        Write-Info "Your local report has still been saved."
    }
}

#endregion

#region Main

Write-Header

Ensure-PnPModule

if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath | Out-Null
}

$largeFileBytes = [long]$LargeFileMB * 1MB

if (Test-AppRegAuth) {
    Write-Step "Authenticating via app registration (TenantId: $TenantId, ClientId: $ClientId)..."
} else {
    Write-Step "Authenticating with Microsoft 365 (browser login will open)..."
}

$allFindings = [System.Collections.Generic.List[hashtable]]::new()

if ($SiteUrl -eq "all") {
    try {
        $resolvedAdminUrl = if ($AdminUrl -ne "") {
            $AdminUrl
        } else {
            Read-Host "Enter your SharePoint Admin Center URL (e.g. https://contoso-admin.sharepoint.com)"
        }
        Connect-ToSite -Url $resolvedAdminUrl
        $siteUrls = Get-AllSiteUrls -AdminUrl $resolvedAdminUrl
        Write-Success "Found $($siteUrls.Count) site collection(s) to scan."
    } catch {
        Write-Host "[!] Could not enumerate sites: $_" -ForegroundColor Red
        Write-Info "Tip: Ensure you have SharePoint Administrator role (or Sites.FullControl.All for app reg)."
        exit 1
    }
} else {
    try {
        Connect-ToSite -Url $SiteUrl
    } catch {
        Write-Host "[!] Authentication failed: $_" -ForegroundColor Red
        exit 1
    }
    $siteUrls = @($SiteUrl)
}

$siteCount = 0
foreach ($url in $siteUrls) {
    $siteCount++
    Write-Step "[$siteCount/$($siteUrls.Count)] Scanning: $url"
    $siteFindings = Get-SiteDriveItems -SiteWebUrl $url -StaleThresholdDays $StaleThresholdDays -LargeFileBytes $largeFileBytes -VersionBloatMultiplier $VersionBloatMultiplier
    foreach ($f in $siteFindings) { $allFindings.Add($f) }
    Write-Info "  Found $($siteFindings.Count) finding(s) on this site."
}

$totalSavings = ($allFindings | Measure-Object -Property PotentialSavingsBytes -Sum).Sum ?? 0

$summary = @{
    TotalFindings      = $allFindings.Count
    TotalSavingsBytes  = [long]$totalSavings
    TotalSavingsGB     = [math]::Round($totalSavings / 1GB, 2)
    EstimatedMonthlySavingsUSD = [math]::Round(($totalSavings / 1GB) * 0.20, 2)
    StaleCount         = ($allFindings | Where-Object { $_.FindingType -eq "stale" }).Count
    LargeFileCount     = ($allFindings | Where-Object { $_.FindingType -eq "large_file" }).Count
    VersionBloatCount  = ($allFindings | Where-Object { $_.FindingType -eq "version_bloat" }).Count
    DuplicateCount     = ($allFindings | Where-Object { $_.FindingType -eq "duplicate" }).Count
}

Write-Host ""
Write-Host "Scan Complete" -ForegroundColor Cyan
Write-Host "  Total findings : $($summary.TotalFindings)"
Write-Host "  Potential savings: $(Format-Bytes $summary.TotalSavingsBytes) (~`$$($summary.EstimatedMonthlySavingsUSD)/mo)"
Write-Host ""

$scope = if ($SiteUrl -eq "all") { "All Sites" } else { $SiteUrl }

$reportPath = New-HtmlReport -Findings $allFindings -Summary $summary -OutputPath $OutputPath -ScannedScope $scope
Write-Success "HTML report saved: $reportPath"

if ($ExportCsv) {
    $csvPath = Export-FindingsCsv -Findings $allFindings -OutputPath $OutputPath
    Write-Success "CSV export saved: $csvPath"
}

if ($ExportJson) {
    $jsonPath = Export-FindingsJson -Findings $allFindings -OutputPath $OutputPath
    Write-Success "JSON export saved: $jsonPath"
}

if ($ApiKey -ne "") {
    Upload-ToStorageScan -Findings $allFindings -Summary $summary -ApiKey $ApiKey -ApiBaseUrl $ApiBaseUrl -ScannedScope $scope
}

Write-Host ""
Write-Host "Opening report in browser..." -ForegroundColor Gray
try {
    Start-Process $reportPath
} catch {
    Write-Info "Could not open browser automatically. Open the report manually: $reportPath"
}

#endregion
