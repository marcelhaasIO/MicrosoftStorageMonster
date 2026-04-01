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
    Write-Host ""
    Write-Host "====================================================" -ForegroundColor Cyan
    Write-Host "  StorageScan Local - SharePoint Storage Analyzer   " -ForegroundColor Cyan
    Write-Host "====================================================" -ForegroundColor Cyan
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
        $badgeColor = switch ($_.FindingType) {
            "stale"        { "#f59e0b" }
            "large_file"   { "#3b82f6" }
            "version_bloat"{ "#8b5cf6" }
            "duplicate"    { "#ef4444" }
            default        { "#6b7280" }
        }
        $label = switch ($_.FindingType) {
            "stale"        { "Stale" }
            "large_file"   { "Large File" }
            "version_bloat"{ "Version Bloat" }
            "duplicate"    { "Duplicate" }
            default        { $_.FindingType }
        }
        $savings = Format-Bytes $_.PotentialSavingsBytes
        $size = Format-Bytes $_.FileSizeBytes
        $modDate = if ($_.LastModified) { $_.LastModified.Substring(0,10) } else { "Unknown" }

        "<tr>
            <td><span style='background:$badgeColor;color:#fff;padding:2px 8px;border-radius:9999px;font-size:11px;font-weight:600'>$label</span></td>
            <td title='$($_.FilePath)'>$($_.FileName)</td>
            <td style='color:#9ca3af;font-size:12px' title='$($_.FilePath)'>$($_.SiteUrl -replace 'https://[^/]+','')</td>
            <td>$size</td>
            <td>$modDate</td>
            <td>$($_.VersionsCount)</td>
            <td style='color:#22c55e;font-weight:600'>$savings</td>
            <td style='color:#6b7280;font-size:12px'>$($_.Details)</td>
        </tr>"
    }) -join "`n"

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>StorageScan Local Report - $now</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#0f0f12;color:#e5e7eb;min-height:100vh}
    .header{background:linear-gradient(135deg,#1a1a22 0%,#12121a 100%);border-bottom:1px solid #27272a;padding:32px 40px}
    .header h1{font-size:28px;font-weight:800;color:#fff;letter-spacing:-0.5px}
    .header h1 span{color:#eab308}
    .header p{color:#6b7280;margin-top:6px;font-size:14px}
    .container{max-width:1400px;margin:0 auto;padding:32px 40px}
    .summary-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:16px;margin-bottom:32px}
    .card{background:#18181b;border:1px solid #27272a;border-radius:12px;padding:20px}
    .card .label{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:0.08em;color:#6b7280;margin-bottom:8px}
    .card .value{font-size:26px;font-weight:800;color:#fff}
    .card .value.green{color:#22c55e}
    .card .value.yellow{color:#eab308}
    .card .sub{font-size:12px;color:#6b7280;margin-top:4px}
    .section-title{font-size:18px;font-weight:700;color:#fff;margin-bottom:16px;display:flex;align-items:center;gap:10px}
    .section-title::before{content:'';display:inline-block;width:4px;height:18px;background:#eab308;border-radius:2px}
    .table-wrap{overflow-x:auto;border-radius:12px;border:1px solid #27272a;margin-bottom:32px}
    table{width:100%;border-collapse:collapse;font-size:13px}
    thead{background:#18181b}
    thead th{padding:12px 14px;text-align:left;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:0.06em;color:#6b7280;border-bottom:1px solid #27272a;cursor:pointer;user-select:none;white-space:nowrap}
    thead th:hover{color:#eab308}
    thead th::after{content:' \25B4\25BE';opacity:0.3;font-size:9px}
    tbody tr{border-bottom:1px solid #1f1f23;transition:background 0.15s}
    tbody tr:hover{background:#1f1f23}
    tbody td{padding:10px 14px;color:#d1d5db;vertical-align:middle;max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .search-bar{display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap}
    .search-bar input,.search-bar select{background:#18181b;border:1px solid #27272a;border-radius:8px;color:#e5e7eb;padding:8px 14px;font-size:13px;outline:none}
    .search-bar input{flex:1;min-width:200px}
    .search-bar input:focus,.search-bar select:focus{border-color:#eab308}
    .type-pills{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:16px}
    .pill{padding:5px 14px;border-radius:9999px;font-size:12px;font-weight:600;cursor:pointer;border:1px solid #27272a;background:#18181b;color:#9ca3af;transition:all 0.15s}
    .pill.active,.pill:hover{border-color:#eab308;color:#eab308;background:#1c1a0a}
    .pill.all{border-color:#eab308;color:#eab308;background:#1c1a0a}
    .footer{text-align:center;color:#374151;font-size:12px;padding:32px;border-top:1px solid #18181b}
    .footer a{color:#eab308;text-decoration:none}
    .no-results{text-align:center;padding:48px;color:#4b5563}
    @media(max-width:768px){.container,.header{padding:20px 16px}.card .value{font-size:20px}}
  </style>
</head>
<body>
<div class="header">
  <h1>SharePoint <span>StorageScan</span> &mdash; Local Report</h1>
  <p>Generated: $now &bull; Scope: $ScannedScope &bull; Thresholds: Stale &gt; $StaleThresholdDays days, Large &gt; $LargeFileMB MB, Version Bloat &gt; ${VersionBloatMultiplier}x</p>
</div>
<div class="container">
  <div class="summary-grid">
    <div class="card">
      <div class="label">Total Findings</div>
      <div class="value yellow">$($Findings.Count)</div>
      <div class="sub">Across all categories</div>
    </div>
    <div class="card">
      <div class="label">Potential Savings</div>
      <div class="value green">$gbSavings GB</div>
      <div class="sub">&asymp; `$$costSavings / month @ `$0.20/GB</div>
    </div>
    <div class="card">
      <div class="label">Stale Files</div>
      <div class="value">$($countByType.stale)</div>
      <div class="sub">$(Format-Bytes $savingsByType.stale) recoverable</div>
    </div>
    <div class="card">
      <div class="label">Large Files</div>
      <div class="value">$($countByType.large_file)</div>
      <div class="sub">$(Format-Bytes $savingsByType.large_file) total</div>
    </div>
    <div class="card">
      <div class="label">Version Bloat</div>
      <div class="value">$($countByType.version_bloat)</div>
      <div class="sub">$(Format-Bytes $savingsByType.version_bloat) recoverable</div>
    </div>
    <div class="card">
      <div class="label">Duplicates</div>
      <div class="value">$($countByType.duplicate)</div>
      <div class="sub">$(Format-Bytes $savingsByType.duplicate) recoverable</div>
    </div>
  </div>

  <div class="section-title">Findings</div>
  <div class="search-bar">
    <input type="text" id="searchInput" placeholder="Filter by file name or path..." oninput="filterTable()">
    <select id="typeFilter" onchange="filterTable()">
      <option value="">All Types</option>
      <option value="stale">Stale</option>
      <option value="large_file">Large File</option>
      <option value="version_bloat">Version Bloat</option>
      <option value="duplicate">Duplicate</option>
    </select>
  </div>
  <div class="table-wrap">
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
  Generated by <a href="https://storagescan.app" target="_blank">StorageScan Local</a> &mdash; Your data never left your machine.
</div>
<script>
  let sortDir = {};
  function filterTable() {
    const q = document.getElementById('searchInput').value.toLowerCase();
    const type = document.getElementById('typeFilter').value;
    const rows = document.querySelectorAll('#tableBody tr');
    let visible = 0;
    rows.forEach(r => {
      const text = r.innerText.toLowerCase();
      const badge = r.querySelector('span') ? r.querySelector('span').innerText.toLowerCase().replace(' ','_') : '';
      const show = (!q || text.includes(q)) && (!type || badge.includes(type.replace('_','').toLowerCase()) || r.cells[0].innerText.toLowerCase().replace(' ','_').includes(type));
      r.style.display = show ? '' : 'none';
      if (show) visible++;
    });
    document.getElementById('noResults').style.display = visible === 0 ? 'block' : 'none';
  }
  function sortTable(col) {
    const tbody = document.getElementById('tableBody');
    const rows = Array.from(tbody.querySelectorAll('tr'));
    const asc = !sortDir[col];
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
