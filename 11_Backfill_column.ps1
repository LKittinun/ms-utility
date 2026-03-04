$w          = 55
$border     = "=" * $w
$rule       = "-" * $w
$prohibited = @("blank", "raw_summary", "prtc", "sst", "column_usage_history")

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "  [11]  Backfill existing column" -ForegroundColor Cyan
Write-Host "        Renames column folder with date prefix," -ForegroundColor DarkCyan
Write-Host "        generates project_info.json and column_log.csv" -ForegroundColor DarkCyan
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
    Write-Host ""
    return
}

# -- Confirm ------------------------------------------------------------------
$cItems = @("Run", "Back to main menu")
$cSel   = 0
$cTop   = [Console]::CursorTop
[Console]::SetCursorPosition(0, $cTop)
Write-Host ("  > " + $cItems[0]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
[Console]::SetCursorPosition(0, $cTop + 1)
Write-Host ("    " + $cItems[1]).PadRight($w + 4) -ForegroundColor DarkCyan -NoNewline
[Console]::SetCursorPosition(0, $cTop + 2)
:confirmLoop while ($true) {
    $ck = [Console]::ReadKey($true)
    if ($ck.Key -eq [ConsoleKey]::UpArrow -or $ck.Key -eq [ConsoleKey]::DownArrow) {
        $p = $cSel; $cSel = 1 - $cSel
        [Console]::SetCursorPosition(0, $cTop + $p)
        Write-Host ("    " + $cItems[$p]).PadRight($w + 4) -ForegroundColor $(if ($p -eq 0) { "Cyan" } else { "DarkCyan" }) -NoNewline
        [Console]::SetCursorPosition(0, $cTop + $cSel)
        Write-Host ("  > " + $cItems[$cSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
    } elseif ($ck.Key -eq [ConsoleKey]::Enter) {
        if ($cSel -eq 1) { Clear-Host; .\Main.ps1; return }
        break confirmLoop
    } elseif ($ck.Key -eq [ConsoleKey]::Escape) {
        Clear-Host; .\Main.ps1; return
    }
}
Write-Host ""

# ── Root ──────────────────────────────────────────────────────────────────────
$root = Read-Host "  Root directory (leave blank for Z:\Proteomics)"
if ($root -eq "") { $root = "Z:\Proteomics" }
$projectsRoot = Join-Path $root "Projects"

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

# ── Check if column folder needs a date prefix ────────────────────────────────
$colFolderName    = Split-Path $analyticsPath -Leaf
$colNeedsRename   = $colFolderName -notmatch '^\d{4}-\d{2}-\d{2}_'
$newAnalyticsPath = $analyticsPath
$newColFolderName = $colFolderName
if ($colNeedsRename) {
    $colCreated       = (Get-Item $analyticsPath).CreationTime.ToString("yyyy-MM-dd")
    $newColFolderName = "${colCreated}_${analyticsCol}"
    $newAnalyticsPath = Join-Path $projectsRoot $newColFolderName
}

# ── Scan folders ──────────────────────────────────────────────────────────────
$projects = Get-ChildItem -Path $analyticsPath -Directory |
    Where-Object { $prohibited -notcontains ($_.Name -replace '^\d{4}-\d{2}-\d{2}_','').ToLower() } |
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
if ($colNeedsRename) {
    Write-Host "  Column folder rename:" -ForegroundColor Cyan
    Write-Host "    $colFolderName" -ForegroundColor DarkGray
    Write-Host "    -> $newColFolderName" -ForegroundColor Yellow
    Write-Host "  $rule" -ForegroundColor DarkCyan
}

$newNo = 1
foreach ($p in $projects) {
    $jsonPath = Join-Path $p.FullName "project_info.json"
    $hasJson  = Test-Path $jsonPath

    $sampleFolders = Get-ChildItem -Path $p.FullName -Directory |
        Where-Object { $prohibited -notcontains ($_.Name -replace '^\d{4}-\d{2}-\d{2}_','').ToLower() } |
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

# ── Rename column folder if needed ────────────────────────────────────────────
if ($colNeedsRename) {
    try {
        Rename-Item -Path $analyticsPath -NewName $newColFolderName -ErrorAction Stop
        $analyticsPath = $newAnalyticsPath
        $logFile       = Join-Path $analyticsPath "column_log.csv"
        Write-Host ""
        Write-Host "  Renamed : $colFolderName -> $newColFolderName" -ForegroundColor Green
    } catch {
        Write-Host ""
        Write-Host "  ERROR renaming column folder: $_" -ForegroundColor Red
        Write-Host "  Continuing with original path." -ForegroundColor DarkYellow
    }
}

# ── Write metadata ────────────────────────────────────────────────────────────
Write-Host ""
$logRows = @()
$newNo   = 1

foreach ($p in $projects) {
    $pPath         = Join-Path $analyticsPath $p.Name
    $jsonPath      = Join-Path $pPath "project_info.json"
    $sampleFolders = Get-ChildItem -Path $pPath -Directory |
        Where-Object { $prohibited -notcontains ($_.Name -replace '^\d{4}-\d{2}-\d{2}_','').ToLower() } |
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
