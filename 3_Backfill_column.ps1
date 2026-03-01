$w      = 55
$border = "=" * $w
$rule   = "-" * $w

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "   [3]  Backfill existing column" -ForegroundColor Cyan
Write-Host "        Generates project_info.json and" -ForegroundColor DarkCyan
Write-Host "        column_log.csv for existing folders" -ForegroundColor DarkCyan
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host ""

# ── Root ──────────────────────────────────────────────────────────────────────
$root = Read-Host "  Root directory (leave blank for Z:\Proteomics)"
if ($root -eq "") { $root = "Z:\Proteomics" }

# ── Analytics column ──────────────────────────────────────────────────────────
Write-Host ""
$analyticsCol = Read-Host "  Analytics column number (e.g. C20531700)"
if ($analyticsCol -eq "") {
    Write-Host "  Analytics column number cannot be empty." -ForegroundColor Red
    Write-Host ""
    Write-Host "  $rule" -ForegroundColor DarkCyan
    $nItems = @("Back to main menu", "Exit"); $nSel = 0
    $nTop = [Console]::CursorTop
    [Console]::SetCursorPosition(0, $nTop);     Write-Host ("  > " + $nItems[0]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
    [Console]::SetCursorPosition(0, $nTop + 1); Write-Host ("    " + $nItems[1]).PadRight($w + 4) -ForegroundColor DarkYellow -NoNewline
    [Console]::SetCursorPosition(0, $nTop + 2)
    while ($true) {
        $k = [Console]::ReadKey($true)
        if ($k.Key -eq [ConsoleKey]::UpArrow -or $k.Key -eq [ConsoleKey]::DownArrow) {
            $p = $nSel; $nSel = 1 - $nSel
            [Console]::SetCursorPosition(0, $nTop + $p);    Write-Host ("    " + $nItems[$p]).PadRight($w + 4) -ForegroundColor $(if ($p -eq 0) { "Cyan" } else { "DarkYellow" }) -NoNewline
            [Console]::SetCursorPosition(0, $nTop + $nSel); Write-Host ("  > " + $nItems[$nSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
        } elseif ($k.Key -eq [ConsoleKey]::Enter -or $k.Key -eq [ConsoleKey]::Escape) {
            if ($k.Key -ne [ConsoleKey]::Escape -and $nSel -eq 0) { Clear-Host; .\Main.ps1 }
            else { [Console]::SetCursorPosition(0, $nTop + 3); Write-Host "  Exiting..." -ForegroundColor DarkYellow }
            return
        }
    }
    return
}

$analyticsPath = Join-Path $root $analyticsCol
$logFile       = Join-Path $analyticsPath "column_log.csv"

if (-not (Test-Path $analyticsPath)) {
    Write-Host "  Path not found: $analyticsPath" -ForegroundColor Red
    Write-Host ""
    Write-Host "  $rule" -ForegroundColor DarkCyan
    $nItems = @("Back to main menu", "Exit"); $nSel = 0
    $nTop = [Console]::CursorTop
    [Console]::SetCursorPosition(0, $nTop);     Write-Host ("  > " + $nItems[0]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
    [Console]::SetCursorPosition(0, $nTop + 1); Write-Host ("    " + $nItems[1]).PadRight($w + 4) -ForegroundColor DarkYellow -NoNewline
    [Console]::SetCursorPosition(0, $nTop + 2)
    while ($true) {
        $k = [Console]::ReadKey($true)
        if ($k.Key -eq [ConsoleKey]::UpArrow -or $k.Key -eq [ConsoleKey]::DownArrow) {
            $p = $nSel; $nSel = 1 - $nSel
            [Console]::SetCursorPosition(0, $nTop + $p);    Write-Host ("    " + $nItems[$p]).PadRight($w + 4) -ForegroundColor $(if ($p -eq 0) { "Cyan" } else { "DarkYellow" }) -NoNewline
            [Console]::SetCursorPosition(0, $nTop + $nSel); Write-Host ("  > " + $nItems[$nSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
        } elseif ($k.Key -eq [ConsoleKey]::Enter -or $k.Key -eq [ConsoleKey]::Escape) {
            if ($k.Key -ne [ConsoleKey]::Escape -and $nSel -eq 0) { Clear-Host; .\Main.ps1 }
            else { [Console]::SetCursorPosition(0, $nTop + 3); Write-Host "  Exiting..." -ForegroundColor DarkYellow }
            return
        }
    }
    return
}

# ── Scan folders ──────────────────────────────────────────────────────────────
$skip     = @("Column_usage_history")
$projects = Get-ChildItem -Path $analyticsPath -Directory |
    Where-Object { $skip -notcontains $_.Name } |
    Sort-Object CreationTime

if ($projects.Count -eq 0) {
    Write-Host "  No project folders found in: $analyticsPath" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  $rule" -ForegroundColor DarkCyan
    $nItems = @("Back to main menu", "Exit"); $nSel = 0
    $nTop = [Console]::CursorTop
    [Console]::SetCursorPosition(0, $nTop);     Write-Host ("  > " + $nItems[0]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
    [Console]::SetCursorPosition(0, $nTop + 1); Write-Host ("    " + $nItems[1]).PadRight($w + 4) -ForegroundColor DarkYellow -NoNewline
    [Console]::SetCursorPosition(0, $nTop + 2)
    while ($true) {
        $k = [Console]::ReadKey($true)
        if ($k.Key -eq [ConsoleKey]::UpArrow -or $k.Key -eq [ConsoleKey]::DownArrow) {
            $p = $nSel; $nSel = 1 - $nSel
            [Console]::SetCursorPosition(0, $nTop + $p);    Write-Host ("    " + $nItems[$p]).PadRight($w + 4) -ForegroundColor $(if ($p -eq 0) { "Cyan" } else { "DarkYellow" }) -NoNewline
            [Console]::SetCursorPosition(0, $nTop + $nSel); Write-Host ("  > " + $nItems[$nSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
        } elseif ($k.Key -eq [ConsoleKey]::Enter -or $k.Key -eq [ConsoleKey]::Escape) {
            if ($k.Key -ne [ConsoleKey]::Escape -and $nSel -eq 0) { Clear-Host; .\Main.ps1 }
            else { [Console]::SetCursorPosition(0, $nTop + 3); Write-Host "  Exiting..." -ForegroundColor DarkYellow }
            return
        }
    }
    return
}

# ── Preview ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "  Preview  ($($projects.Count) folders  |  sorted by creation date)" -ForegroundColor Cyan
Write-Host "  $rule" -ForegroundColor DarkCyan

$newNo = 1
foreach ($p in $projects) {
    $jsonPath = Join-Path $p.FullName "project_info.json"
    $hasJson  = Test-Path $jsonPath

    $sampleFolders = Get-ChildItem -Path $p.FullName -Directory |
        Where-Object { $skip -notcontains $_.Name } |
        Select-Object -ExpandProperty Name

    if ($hasJson) {
        $existing  = Get-Content $jsonPath -Raw | ConvertFrom-Json
        $idDisplay = if ($existing.ProjectID) { $existing.ProjectID } else { "(will generate)" }
        Write-Host ("  [$newNo]  " + $p.Name) -ForegroundColor Gray
        Write-Host ("        ID : $idDisplay  (existing - preserved)") -ForegroundColor DarkGray
    } else {
        Write-Host ("  [$newNo]  " + $p.Name) -ForegroundColor White
        Write-Host ("        ID : (will generate)") -ForegroundColor Cyan
    }
    Write-Host ("        Subfolders : " + $(if ($sampleFolders) { $sampleFolders -join ", " } else { "(none)" })) -ForegroundColor DarkGray
    $newNo++
}
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  Gray = already has metadata (ID preserved)" -ForegroundColor DarkGray
Write-Host "  White = new metadata will be generated" -ForegroundColor White
Write-Host ""

# ── Confirm ───────────────────────────────────────────────────────────────────
$cItems = @("Yes, generate metadata", "No, cancel")
$cSel   = 0
$cTop   = [Console]::CursorTop
[Console]::SetCursorPosition(0, $cTop);     Write-Host ("  > " + $cItems[0]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
[Console]::SetCursorPosition(0, $cTop + 1); Write-Host ("    " + $cItems[1]).PadRight($w + 4) -ForegroundColor DarkYellow -NoNewline
[Console]::SetCursorPosition(0, $cTop + 2)
while ($true) {
    $k = [Console]::ReadKey($true)
    if ($k.Key -eq [ConsoleKey]::UpArrow -or $k.Key -eq [ConsoleKey]::DownArrow) {
        $p = $cSel; $cSel = 1 - $cSel
        [Console]::SetCursorPosition(0, $cTop + $p);    Write-Host ("    " + $cItems[$p]).PadRight($w + 4) -ForegroundColor $(if ($p -eq 0) { "Cyan" } else { "DarkYellow" }) -NoNewline
        [Console]::SetCursorPosition(0, $cTop + $cSel); Write-Host ("  > " + $cItems[$cSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
    } elseif ($k.Key -eq [ConsoleKey]::Enter -or $k.Key -eq [ConsoleKey]::Escape) {
        [Console]::SetCursorPosition(0, $cTop + 2)
        if ($k.Key -ne [ConsoleKey]::Escape -and $cSel -eq 0) { break }
        Write-Host "  Cancelled." -ForegroundColor DarkYellow
        Write-Host ""
        Write-Host "  $rule" -ForegroundColor DarkCyan
        $nItems = @("Back to main menu", "Exit"); $nSel = 0
        $nTop = [Console]::CursorTop
        [Console]::SetCursorPosition(0, $nTop);     Write-Host ("  > " + $nItems[0]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
        [Console]::SetCursorPosition(0, $nTop + 1); Write-Host ("    " + $nItems[1]).PadRight($w + 4) -ForegroundColor DarkYellow -NoNewline
        [Console]::SetCursorPosition(0, $nTop + 2)
        while ($true) {
            $k2 = [Console]::ReadKey($true)
            if ($k2.Key -eq [ConsoleKey]::UpArrow -or $k2.Key -eq [ConsoleKey]::DownArrow) {
                $p = $nSel; $nSel = 1 - $nSel
                [Console]::SetCursorPosition(0, $nTop + $p);    Write-Host ("    " + $nItems[$p]).PadRight($w + 4) -ForegroundColor $(if ($p -eq 0) { "Cyan" } else { "DarkYellow" }) -NoNewline
                [Console]::SetCursorPosition(0, $nTop + $nSel); Write-Host ("  > " + $nItems[$nSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
            } elseif ($k2.Key -eq [ConsoleKey]::Enter -or $k2.Key -eq [ConsoleKey]::Escape) {
                if ($k2.Key -ne [ConsoleKey]::Escape -and $nSel -eq 0) { Clear-Host; .\Main.ps1 }
                else { [Console]::SetCursorPosition(0, $nTop + 3); Write-Host "  Exiting..." -ForegroundColor DarkYellow }
                return
            }
        }
        return
    }
}

# ── Write metadata ────────────────────────────────────────────────────────────
Write-Host ""
$logRows = @()
$newNo   = 1

foreach ($p in $projects) {
    $jsonPath      = Join-Path $p.FullName "project_info.json"
    $sampleFolders = Get-ChildItem -Path $p.FullName -Directory |
        Where-Object { $skip -notcontains $_.Name } |
        Select-Object -ExpandProperty Name

    if (Test-Path $jsonPath) {
        $existing  = Get-Content $jsonPath -Raw | ConvertFrom-Json
        $projectID = if ($existing.ProjectID) { $existing.ProjectID }
                     else { -join ((65..90) + (48..57) | Get-Random -Count 8 | ForEach-Object { [char]$_ }) }
        Write-Host "  [$newNo] $($p.Name)  (preserved)" -ForegroundColor Gray
    } else {
        $projectID = -join ((65..90) + (48..57) | Get-Random -Count 8 | ForEach-Object { [char]$_ })
        Write-Host "  [$newNo] $($p.Name)  ID: $projectID" -ForegroundColor Green
    }

    $created = $p.CreationTime.ToString("yyyy-MM-dd HH:mm")

    [PSCustomObject]@{
        ProjectID       = $projectID
        Project         = $p.Name
        PI              = if ((Test-Path $jsonPath) -and $existing.PI) { $existing.PI } else { $null }
        AnalyticsColumn = $analyticsCol
        TrapColumn      = $null
        ProjectNo       = $newNo
        Created         = $created
        SampleFolders   = @($sampleFolders)
    } | ConvertTo-Json | Out-File -FilePath $jsonPath -Encoding UTF8

    $logRows += [PSCustomObject]@{
        ProjectID       = $projectID
        ProjectNo       = $newNo
        Date            = $created
        Project         = $p.Name
        PI              = if ((Test-Path $jsonPath) -and $existing.PI) { $existing.PI } else { "" }
        AnalyticsColumn = $analyticsCol
        TrapColumn      = ""
        SampleFolders   = $sampleFolders -join ";"
    }

    $newNo++
}

$logRows | Export-Csv $logFile -NoTypeInformation
Write-Host ""
Write-Host "  Rebuilt : $logFile" -ForegroundColor Green

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "  Done!  $($projects.Count) project(s) processed." -ForegroundColor Cyan
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  TrapColumn is blank for backfilled projects." -ForegroundColor DarkGray
Write-Host "  Use [7] Repair project order to fix ordering" -ForegroundColor DarkGray
Write-Host "  or edit project_info.json files directly." -ForegroundColor DarkGray
Write-Host "  $border" -ForegroundColor DarkCyan

# ── Navigation ────────────────────────────────────────────────────────────────
$nItems = @("Back to main menu", "Exit")
$nSel   = 0
Write-Host ""
Write-Host "  $rule" -ForegroundColor DarkCyan
$nTop = [Console]::CursorTop
[Console]::SetCursorPosition(0, $nTop)
Write-Host ("  > " + $nItems[0]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
[Console]::SetCursorPosition(0, $nTop + 1)
Write-Host ("    " + $nItems[1]).PadRight($w + 4) -ForegroundColor DarkYellow -NoNewline
[Console]::SetCursorPosition(0, $nTop + 2)
while ($true) {
    $k = [Console]::ReadKey($true)
    if ($k.Key -eq [ConsoleKey]::UpArrow -or $k.Key -eq [ConsoleKey]::DownArrow) {
        $p = $nSel; $nSel = 1 - $nSel
        [Console]::SetCursorPosition(0, $nTop + $p);    Write-Host ("    " + $nItems[$p]).PadRight($w + 4) -ForegroundColor $(if ($p -eq 0) { "Cyan" } else { "DarkYellow" }) -NoNewline
        [Console]::SetCursorPosition(0, $nTop + $nSel); Write-Host ("  > " + $nItems[$nSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
    } elseif ($k.Key -eq [ConsoleKey]::Enter -or $k.Key -eq [ConsoleKey]::Escape) {
        if ($k.Key -ne [ConsoleKey]::Escape -and $nSel -eq 0) { Clear-Host; .\Main.ps1 }
        else { [Console]::SetCursorPosition(0, $nTop + 3); Write-Host "  Exiting..." -ForegroundColor DarkYellow }
        return
    }
}
