$w          = 55
$border     = "=" * $w
$rule       = "-" * $w
$prohibited = @("blank", "raw_summary", "prtc", "sst", "column_usage_history")
. (Join-Path $PSScriptRoot "lib\Menu.ps1")

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "  [12]  Repair project order" -ForegroundColor Cyan
Write-Host "        Re-numbers projects by creation date" -ForegroundColor DarkCyan
Write-Host "        and rebuilds column_log.csv" -ForegroundColor DarkCyan
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host ""

# -- Password -----------------------------------------------------------------
$pwSS    = Read-Host "  Password" -AsSecureString
$pwPlain = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
               [Runtime.InteropServices.Marshal]::SecureStringToBSTR($pwSS))
$pwHash  = [System.BitConverter]::ToString(
               [System.Security.Cryptography.SHA256]::Create().ComputeHash(
                   [System.Text.Encoding]::UTF8.GetBytes($pwPlain)
               )).Replace("-","").ToLower()
if ($pwHash -ne "15e2b0d3c33891ebb0f1ef609ec419420c20e320ce94c65fbc8c3312448eb225") {
    Write-Host "  Access denied." -ForegroundColor Red
    Show-NavExit
    return
}

# -- Confirm ------------------------------------------------------------------
$r = Show-Menu -Items @("Run", "Back to main menu") -AllowEscape
if ($r.Action -ne "select" -or $r.Index -eq 1) { Return-ToMain; return }
Write-Host ""

# ── Root ──────────────────────────────────────────────────────────────────────
$_cfg         = if (Test-Path (Join-Path $PSScriptRoot "config.json")) { Get-Content (Join-Path $PSScriptRoot "config.json") -Raw | ConvertFrom-Json } else { $null }
$root         = if ($_cfg -and $_cfg.Root) { $_cfg.Root } else { "Z:\Proteomics" }
$projectsRoot = Join-Path $root "Projects"
Write-Host "  Root : $root" -ForegroundColor DarkGray

# ── Analytics column ──────────────────────────────────────────────────────────
Write-Host ""
$analyticsCol = Read-Host "  Analytics column number (e.g. C20531700)"
if ($analyticsCol -eq "") {
    Write-Host "  Analytics column number cannot be empty." -ForegroundColor Red
    Write-Host ""
    Write-Host "  $rule" -ForegroundColor DarkCyan
    Show-NavExit
    return
}

# ── Resolve analytics column folder (date-prefixed) ──────────────────────────
$analyticsPath = $null
if (Test-Path $projectsRoot) {
    $existingColDir = Get-ChildItem $projectsRoot -Directory |
        Where-Object { $_.Name -like "*_$analyticsCol" } |
        Select-Object -First 1
    if ($existingColDir) { $analyticsPath = $existingColDir.FullName }
}
if (-not $analyticsPath) { $analyticsPath = Join-Path $projectsRoot $analyticsCol }
$logFile = Join-Path $analyticsPath "column_log.csv"

if (-not (Test-Path $analyticsPath)) {
    Write-Host "  Path not found: $analyticsPath" -ForegroundColor Red
    Write-Host ""
    Write-Host "  $rule" -ForegroundColor DarkCyan
    Show-NavExit
    return
}

# ── Scan projects ─────────────────────────────────────────────────────────────
$projects = Get-ChildItem -Path $analyticsPath -Directory |
    Where-Object { $prohibited -notcontains ($_.Name -replace '^\d{4}-\d{2}-\d{2}_','').ToLower() } |
    Where-Object { Test-Path (Join-Path $_.FullName "project_info.json") } |
    ForEach-Object {
        $f    = $_
        $json = Get-Content (Join-Path $f.FullName "project_info.json") -Raw | ConvertFrom-Json
        $createdDt = try {
            [datetime]::ParseExact($json.Created, "yyyy-MM-dd HH:mm", $null)
        } catch {
            Write-Warning "  Cannot parse Created date '$($json.Created)' for '$($f.Name)' - using folder date"
            $f.CreationTime
        }
        [PSCustomObject]@{
            Folder    = $f.FullName
            Name      = $f.Name
            Created   = $createdDt
            OldNo     = $json.ProjectNo
            Json      = $json
        }
    } |
    Sort-Object Created

if ($projects.Count -eq 0) {
    Write-Host "  No projects with project_info.json found in: $analyticsPath" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  $rule" -ForegroundColor DarkCyan
    Show-NavExit
    return
}

# ── Preview ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "  Proposed re-numbering (sorted by creation date):" -ForegroundColor Cyan
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host ("  " + "No.".PadRight(6) + "Was".PadRight(6) + "Created".PadRight(18) + "Project") -ForegroundColor DarkCyan

$newNo = 1
foreach ($p in $projects) {
    $changed = $p.OldNo -ne $newNo
    $color   = if ($changed) { "Yellow" } else { "White" }
    $marker  = if ($changed) { " *" } else { "" }
    Write-Host ("  " + "$newNo".PadRight(6) + "$($p.OldNo)".PadRight(6) + $p.Created.ToString("yyyy-MM-dd HH:mm").PadRight(18) + $p.Name + $marker) -ForegroundColor $color
    $newNo++
}
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  * = number will change" -ForegroundColor DarkCyan
Write-Host ""

# ── Confirm ───────────────────────────────────────────────────────────────────
$cRes = Show-Menu -Items @("Yes, apply changes", "No, cancel") -AllowEscape
if ($cRes.Action -ne "select" -or $cRes.Index -eq 1) {
    Write-Host "  Cancelled." -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "  $rule" -ForegroundColor DarkCyan
    Show-NavExit
    return
}

# ── Backup ────────────────────────────────────────────────────────────────────
# Copies each data file to <Root>\Backup\<Script>_<yyyyMMdd_HHmmss>\<path relative
# to Projects> once per run, before it is overwritten. Folder is created lazily.
$backupBase = $projectsRoot
$backupDir  = Join-Path (Join-Path $root "Backup") ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath) + "_" + (Get-Date -Format "yyyyMMdd_HHmmss"))
$backupDone = @{}
function Backup-DataFile ([string]$path) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $true }
    $full = [System.IO.Path]::GetFullPath($path)
    $key  = $full.ToLower()
    if ($backupDone.ContainsKey($key)) { return $true }
    $base = [System.IO.Path]::GetFullPath($backupBase).TrimEnd("\")
    if ($full.StartsWith($base + "\", [System.StringComparison]::OrdinalIgnoreCase)) {
        $rel = $full.Substring($base.Length + 1)
    } else {
        $rel = ($full -replace '^([A-Za-z]):', '$1').TrimStart("\")
    }
    $dest = Join-Path $backupDir $rel
    try {
        [System.IO.Directory]::CreateDirectory((Split-Path $dest -Parent)) | Out-Null
        Copy-Item -LiteralPath $full -Destination $dest -Force -ErrorAction Stop
    } catch {
        Write-Host "  [ERROR] Backup failed: $full - $_" -ForegroundColor Red
        return $false
    }
    $backupDone[$key] = $true
    return $true
}

# Back up every file this run will overwrite before changing anything
$backupOk = $true
foreach ($p in $projects) {
    if (-not (Backup-DataFile (Join-Path $p.Folder "project_info.json"))) { $backupOk = $false; break }
}
if ($backupOk) { $backupOk = Backup-DataFile $logFile }
if (-not $backupOk) {
    Write-Host "  No changes applied." -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "  $rule" -ForegroundColor DarkCyan
    Show-NavExit
    return
}

# ── Apply ─────────────────────────────────────────────────────────────────────
Write-Host ""
$newLogRows = @()
$newNo = 1
foreach ($p in $projects) {
    # Generate ProjectID before writing JSON so it is persisted on disk
    if (-not $p.Json.ProjectID) {
        $p.Json | Add-Member -NotePropertyName ProjectID -NotePropertyValue (
            -join ((65..90) + (48..57) | Get-Random -Count 8 | ForEach-Object { [char]$_ })
        ) -Force
    }
    $p.Json.ProjectNo = $newNo
    $p.Json | ConvertTo-Json | Out-File -FilePath (Join-Path $p.Folder "project_info.json") -Encoding UTF8
    $changed = if ($p.OldNo -ne $newNo) { " (was $($p.OldNo))" } else { "" }
    Write-Host "  [$newNo] $($p.Name)$changed" -ForegroundColor $(if ($p.OldNo -ne $newNo) { "Yellow" } else { "Green" })

    $newLogRows += [PSCustomObject]@{
        ProjectID             = $p.Json.ProjectID
        ProjectNo             = $newNo
        Date                  = $p.Json.Created
        Project               = $p.Json.Project
        PI                    = if ($null -eq $p.Json.PI) { "" } else { $p.Json.PI }
        AnalyticsColumn       = $p.Json.AnalyticsColumn
        ColumnDescription     = if ($null -eq $p.Json.ColumnDescription) { "" } else { $p.Json.ColumnDescription }
        TrapColumn            = if ($null -eq $p.Json.TrapColumn) { "" } else { $p.Json.TrapColumn }
        TrapColumnDescription = if ($null -eq $p.Json.TrapColumnDescription) { "" } else { $p.Json.TrapColumnDescription }
        SampleFolders         = ($p.Json.SampleFolders -join ";")
    }
    $newNo++
}

# Rebuild column_log.csv from scratch
$newLogRows | Export-Csv $logFile -NoTypeInformation
Write-Host ""
Write-Host "  Rebuilt : $logFile" -ForegroundColor Green
if ($backupDone.Count -gt 0) { Write-Host "  Backup : $backupDir" -ForegroundColor DarkGray }

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "  Done!  $($projects.Count) project(s) re-numbered." -ForegroundColor Cyan
Write-Host "  $border" -ForegroundColor DarkCyan

# ── Navigation ────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  $rule" -ForegroundColor DarkCyan
Show-NavExit
return
