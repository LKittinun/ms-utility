$w      = 55
$border = "=" * $w
$rule   = "-" * $w

. (Join-Path $PSScriptRoot "lib\Menu.ps1")
. (Join-Path $PSScriptRoot "lib\Names.ps1")

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "   [15] Data health check" -ForegroundColor Cyan
Write-Host "        Read-only scan for data problems" -ForegroundColor DarkCyan
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host ""

# -- Config -------------------------------------------------------------------
$configFile = Join-Path $PSScriptRoot "config.json"
$_cfg       = if (Test-Path $configFile) { Get-Content $configFile -Raw | ConvertFrom-Json } else { $null }
$_rootBase  = if ($_cfg -and $_cfg.Root) { $_cfg.Root } else { "Z:\Proteomics" }
$root       = Join-Path $_rootBase "Projects"
Write-Host "  Root : $_rootBase" -ForegroundColor DarkGray
Write-Host ""

# -- Confirm ------------------------------------------------------------------
$r = Show-Menu -Items @("Run", "Back to main menu") -AllowEscape
if ($r.Action -ne "select" -or $r.Index -eq 1) { Return-ToMain; return }
Write-Host ""

if (-not (Test-Path $root)) {
    Write-Host "  [ERROR] Projects root not found: $root" -ForegroundColor Red
    Show-NavExit
    return
}

# -- Scan ---------------------------------------------------------------------
Write-Host "  Scanning $root ..." -ForegroundColor DarkCyan
$projects = [System.Collections.Generic.List[object]]::new()
$columns  = [System.Collections.Generic.List[object]]::new()
foreach ($colDir in Get-ChildItem $root -Directory -ErrorAction SilentlyContinue) {
    $colProjects = @()
    foreach ($pDir in Get-ChildItem $colDir.FullName -Directory -ErrorAction SilentlyContinue) {
        $jp = Join-Path $pDir.FullName "project_info.json"
        if (-not (Test-Path $jp)) { continue }
        $info = $null
        try { $info = Get-Content $jp -Raw | ConvertFrom-Json } catch {}
        $p = [PSCustomObject]@{ Column = $colDir.Name; Folder = $pDir.Name; Path = $pDir.FullName; Info = $info }
        $projects.Add($p)
        $colProjects += $p
    }
    $logFile = Join-Path $colDir.FullName "column_log.csv"
    $logRows = @()
    if (Test-Path $logFile) { try { $logRows = @(Import-Csv $logFile) } catch {} }
    $columns.Add([PSCustomObject]@{
        Name     = $colDir.Name
        Path     = $colDir.FullName
        HasInfo  = Test-Path (Join-Path $colDir.FullName "column_info.json")
        HasLog   = Test-Path $logFile
        LogRows  = $logRows
        Projects = $colProjects
    })
}
Write-Host "  Found $($projects.Count) project(s) in $($columns.Count) column folder(s)." -ForegroundColor DarkGray

# Each section: title + list of finding lines
$sections = [System.Collections.Generic.List[object]]::new()
function Add-Section([string]$title, [string]$advice, $lines) {
    $sections.Add([PSCustomObject]@{ Title = $title; Advice = $advice; Lines = @($lines) })
}
function Label($p) { return "$($p.Column)\$($p.Folder)" }

# 1. Unreadable project_info.json
Add-Section "Unreadable project_info.json" "Open the file and fix the JSON by hand (or restore from Backup)." @(
    $projects | Where-Object { $null -eq $_.Info } | ForEach-Object { Label $_ }
)
$ok = @($projects | Where-Object { $null -ne $_.Info })

# 2. Missing PI
Add-Section "Projects with no PI" "Fill in with [1] (Edit a field) or [14] Sync from overview CSV." @(
    $ok | Where-Object { -not "$($_.Info.PI)".Trim() } | ForEach-Object { Label $_ }
)

# 3. Possible duplicate PIs (same person, different spelling)
$piNames = @($ok | ForEach-Object { "$($_.Info.PI)" } | Where-Object { $_.Trim() } | Sort-Object -Unique)
$piCount = @{}
foreach ($p in $ok) { $n = "$($p.Info.PI)"; if ($n.Trim()) { $piCount[$n] = 1 + [int]$piCount[$n] } }
$dupLines = @()
for ($i = 0; $i -lt $piNames.Count; $i++) {
    for ($j = $i + 1; $j -lt $piNames.Count; $j++) {
        $a = $piNames[$i]; $b = $piNames[$j]
        $same = (Get-NameKey $a) -eq (Get-NameKey $b)
        if (-not $same) { $same = [bool](Find-SimilarPI $a @($b)) -or [bool](Find-SimilarPI $b @($a)) }
        if ($same) { $dupLines += "'$a' ($($piCount[$a])) ~ '$b' ($($piCount[$b]))" }
    }
}
Add-Section "Possible duplicate PIs" "If they are the same person, correct the spelling in the overview CSV and run [14]." $dupLines

# 4. Odd project names
Add-Section "Odd project names" "Project names should not repeat the date or start/end with '_' or spaces." @(
    $ok | ForEach-Object {
        $n = "$($_.Info.Project)"
        $why = @()
        if ($n -match '^\d{4}-\d{2}-\d{2}_')     { $why += "contains date prefix" }
        if ($n -match '^[_\s]|[_\s]$')           { $why += "leading/trailing _ or space" }
        if ($n -match '__')                      { $why += "double underscore" }
        if ($n -eq "")                           { $why += "empty" }
        if ($why.Count -gt 0) { "$(Label $_)  ->  '$n'  ($($why -join ', '))" }
    }
)

# 5. Folder name does not match <date>_<Project>
# (older projects have no date prefix - folder == Project is fine too)
Add-Section "Folder name does not match project name" "Expected folder 'YYYY-MM-DD_<Project>' (or '<Project>' for older projects)." @(
    $ok | Where-Object { $_.Folder -notmatch ('^(\d{4}-\d{2}-\d{2}_)?' + [regex]::Escape("$($_.Info.Project)") + '$') } |
        ForEach-Object { "$(Label $_)  (Project = '$($_.Info.Project)')" }
)

# 6. Sample folders listed but missing on disk
Add-Section "Sample folders missing on disk" "Re-run [1] on the project (it re-creates folders) or remove them from the list." @(
    $ok | ForEach-Object {
        $p = $_
        foreach ($sf in @($p.Info.SampleFolders)) {
            if ($sf -and -not (Test-Path (Join-Path $p.Path $sf))) { "$(Label $p)\$sf" }
        }
    }
)

# 7. Duplicate ProjectIDs
Add-Section "Duplicate ProjectIDs" "Each project needs a unique ID - fix by hand." @(
    $ok | Where-Object { $_.Info.ProjectID } | Group-Object { $_.Info.ProjectID } | Where-Object { $_.Count -gt 1 } |
        ForEach-Object { "$($_.Name): $(($_.Group | ForEach-Object { Label $_ }) -join ', ')" }
)

# 8. ProjectNo duplicates / gaps per column
Add-Section "Project numbering problems" "Run [12] Repair project order." @(
    $columns | ForEach-Object {
        $c = $_
        $nos = @($c.Projects | Where-Object { $_.Info } | ForEach-Object { [int]("0" + $_.Info.ProjectNo) } | Sort-Object)
        if ($nos.Count -gt 0) {
            $dups = @($nos | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
            $gaps = @(1..$nos.Count | Where-Object { $nos -notcontains $_ })
            if ($dups.Count -gt 0) { "$($c.Name): duplicate no. $($dups -join ', ')" }
            if ($gaps.Count -gt 0) { "$($c.Name): missing no. $($gaps -join ', ')" }
        }
    }
)

# 9. Column log vs project files
$logLines = @()
foreach ($c in $columns) {
    if ($c.Projects.Count -eq 0) { continue }
    if (-not $c.HasLog) { $logLines += "$($c.Name): column_log.csv missing"; continue }
    $ids   = @($c.Projects | Where-Object { $_.Info } | ForEach-Object { "$($_.Info.ProjectID)" })
    $names = @($c.Projects | Where-Object { $_.Info } | ForEach-Object { "$($_.Info.Project)" })
    foreach ($row in $c.LogRows) {
        if (-not (($row.ProjectID -and $ids -contains $row.ProjectID) -or ($names -contains $row.Project))) {
            $logLines += "$($c.Name): log row '$($row.Project)' has no project folder"
        }
    }
    foreach ($p in $c.Projects | Where-Object { $_.Info }) {
        $row = $c.LogRows | Where-Object { ($_.ProjectID -and $_.ProjectID -eq $p.Info.ProjectID) -or $_.Project -eq $p.Info.Project } | Select-Object -First 1
        if (-not $row) { $logLines += "$($c.Name): project '$($p.Info.Project)' missing from log"; continue }
        $diff = @()
        if ("$($row.PI)" -cne "$($p.Info.PI)") { $diff += "PI" }
        $sfJson = (@($p.Info.SampleFolders) | Where-Object { $_ }) -join ";"
        if ("$($row.SampleFolders)" -ne $sfJson) { $diff += "SampleFolders" }
        if ("$($row.TrapColumn)" -ne "$($p.Info.TrapColumn)") { $diff += "TrapColumn" }
        if ($diff.Count -gt 0) { $logLines += "$($c.Name): '$($p.Info.Project)' differs in $($diff -join ', ')" }
    }
}
Add-Section "Column log out of sync with project files" "Saving any project in that column with [1] rebuilds its log from the project files." $logLines

# 10. Column folders without column_info.json
Add-Section "Column folders without column_info.json" "Run [13] Backfill existing column." @(
    $columns | Where-Object { $_.Projects.Count -gt 0 -and -not $_.HasInfo } | ForEach-Object { $_.Name }
)

# -- Report -------------------------------------------------------------------
$report = [System.Collections.Generic.List[string]]::new()
$report.Add("Data health check - $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
$report.Add("Root: $root   Projects: $($projects.Count)   Columns: $($columns.Count)")
$report.Add("")
$problems = 0
Write-Host ""
foreach ($s in $sections) {
    $n = $s.Lines.Count
    $problems += $n
    if ($n -eq 0) {
        Write-Host "  [ OK ] " -ForegroundColor Green -NoNewline
        Write-Host $s.Title -ForegroundColor Gray
        $report.Add("[ OK ] $($s.Title)")
        continue
    }
    Write-Host "  [ $n ] ".PadRight(9) -ForegroundColor Yellow -NoNewline
    Write-Host $s.Title -ForegroundColor White
    $report.Add("[ $n ] $($s.Title)")
    $shown = 0
    foreach ($line in $s.Lines) {
        $report.Add("        $line")
        if ($shown -lt 8) { Write-Host "         $line" -ForegroundColor DarkGray; $shown++ }
    }
    if ($n -gt 8) { Write-Host "         ... and $($n - 8) more (see saved report)" -ForegroundColor DarkGray }
    Write-Host "         Fix: $($s.Advice)" -ForegroundColor DarkCyan
    $report.Add("        Fix: $($s.Advice)")
}
Write-Host ""
Write-Host "  $rule" -ForegroundColor DarkCyan
if ($problems -eq 0) { Write-Host "  All checks passed." -ForegroundColor Green }
else                 { Write-Host "  $problems finding(s). Nothing was changed." -ForegroundColor Yellow }

# -- Save / navigate ----------------------------------------------------------
Write-Host ""
$r = Show-Menu -Items @("Save report as text file", "Back to main menu", "Exit") -AllowEscape
if ($r.Action -eq "select" -and $r.Index -eq 0) {
    $outFile = Join-Path $_rootBase ("Health_check_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".txt")
    $report | Out-File $outFile -Encoding UTF8
    Write-Host "  Saved: $outFile" -ForegroundColor Green
    Show-NavExit
    return
}
if ($r.Action -eq "select" -and $r.Index -eq 1) { Return-ToMain; return }
Request-Exit
