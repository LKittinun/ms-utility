$w      = 55
$border = "=" * $w
$rule   = "-" * $w

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "   [11] Archive raw files              (to external HDD)" -ForegroundColor Cyan
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

# -- Root check ---------------------------------------------------------------
if (-not (Test-Path $root)) {
    Write-Host "  [ERROR] Projects root not found: $root" -ForegroundColor Red
    Write-Host ""
    pause
    Clear-Host; .\Main.ps1; return
}

# -- Source / destination -----------------------------------------------------
. (Join-Path $PSScriptRoot "lib\Pickers.ps1")
Write-Host "  Source root : $root  (select it to archive all projects)" -ForegroundColor DarkGray
$source = Read-FolderPath "Select the source folder to archive" $root
if ($source -eq "") { Clear-Host; .\Main.ps1; return }
$source = $source.TrimEnd("\")
# Source must be inside the Projects root so the archive mirrors its structure
if (-not ($source -ieq $root.TrimEnd("\") -or $source.StartsWith($root.TrimEnd("\") + "\", [System.StringComparison]::OrdinalIgnoreCase))) {
    Write-Host "  [ERROR] Source must be inside $root" -ForegroundColor Red
    Write-Host ""
    pause
    Clear-Host; .\Main.ps1; return
}
if (-not (Test-Path -LiteralPath $source)) {
    Write-Host "  [ERROR] Path not found: $source" -ForegroundColor Red
    Write-Host ""
    pause
    Clear-Host; .\Main.ps1; return
}
Write-Host ""
$archiveDest = Read-FolderPath "Select the archive destination (e.g. E:\Raw_Archive)" -NoStart
if ($archiveDest -eq "") {
    Write-Host "  Cancelled." -ForegroundColor DarkYellow
    Write-Host ""
    pause
    Clear-Host; .\Main.ps1; return
}

Write-Host ""
Write-Host "  Scanning for .raw files..." -ForegroundColor DarkCyan

# -- Scan ---------------------------------------------------------------------
$rawFiles = @(Get-ChildItem -Path $source -Recurse -Filter "*.raw" -File)

if ($rawFiles.Count -eq 0) {
    Write-Host "  No .raw files found under $source" -ForegroundColor Yellow
    Write-Host ""
    pause
    Clear-Host; .\Main.ps1; return
}

$totalSizeBytes = ($rawFiles | Measure-Object -Property Length -Sum).Sum
$totalSizeGB    = [math]::Round($totalSizeBytes / 1GB, 2)

Write-Host "  Found $($rawFiles.Count) .raw file(s)  ($totalSizeGB GB)" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Source : $source" -ForegroundColor DarkGray
Write-Host "  Dest   : $archiveDest" -ForegroundColor DarkGray
Write-Host ""

# -- Mode picker --------------------------------------------------------------
Write-Host "  Select mode:" -ForegroundColor DarkCyan
$mItems = @("Dry run (list all files, no move)", "Move files (remove from source)", "Backup (copy, keep originals)", "Cancel")
$mSel   = 0
$mTop   = [Console]::CursorTop
for ($mi = 0; $mi -lt $mItems.Count; $mi++) {
    [Console]::SetCursorPosition(0, $mTop + $mi)
    $mText = ("    " + $mItems[$mi]).PadRight($w + 4)
    if ($mi -eq $mSel) { Write-Host $mText -ForegroundColor Black -BackgroundColor Cyan -NoNewline }
    else                { Write-Host $mText -ForegroundColor DarkCyan -NoNewline }
}
[Console]::SetCursorPosition(0, $mTop + $mItems.Count)
:modeLoop while ($true) {
    $mk = [Console]::ReadKey($true)
    if ($mk.Key -eq [ConsoleKey]::UpArrow -or $mk.Key -eq [ConsoleKey]::DownArrow) {
        $mp   = $mSel
        $mSel = if ($mk.Key -eq [ConsoleKey]::UpArrow) { ($mSel - 1 + $mItems.Count) % $mItems.Count } else { ($mSel + 1) % $mItems.Count }
        [Console]::SetCursorPosition(0, $mTop + $mp);   Write-Host ("    " + $mItems[$mp]).PadRight($w + 4) -ForegroundColor DarkCyan -NoNewline
        [Console]::SetCursorPosition(0, $mTop + $mSel); Write-Host ("    " + $mItems[$mSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
    } elseif ($mk.Key -eq [ConsoleKey]::Enter) {
        break modeLoop
    } elseif ($mk.Key -eq [ConsoleKey]::Escape) {
        $mSel = 3
        break modeLoop
    }
}
[Console]::SetCursorPosition(0, $mTop + $mItems.Count)
Write-Host ""

if ($mSel -eq 3) {
    Write-Host "  Cancelled." -ForegroundColor DarkYellow
    Write-Host ""
    pause
    Clear-Host; .\Main.ps1; return
}

# -- Dry run ------------------------------------------------------------------
if ($mSel -eq 0) {
    Write-Host "  $rule" -ForegroundColor DarkCyan
    Write-Host "  Dry run - files that would be moved/copied:" -ForegroundColor Cyan
    Write-Host ""
    $rawFiles | ForEach-Object {
        $rel    = $_.FullName.Substring($source.Length)
        $target = Join-Path $archiveDest $rel
        $sizeMB = [math]::Round($_.Length / 1MB, 1)
        Write-Host "    $($_.Name)  ($sizeMB MB)" -ForegroundColor Gray
        Write-Host "      -> $target" -ForegroundColor DarkGray
    }
    Write-Host ""
    Write-Host "  $rule" -ForegroundColor DarkCyan
    Write-Host "  Total: $($rawFiles.Count) file(s)  $totalSizeGB GB  (no files moved)" -ForegroundColor Cyan
    Write-Host "  $rule" -ForegroundColor DarkCyan

    $nItems = @("Back to main menu", "Exit")
    $nSel   = 0
    Write-Host ""
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
            [Console]::SetCursorPosition(0, $nTop + $nSel); Write-Host ("    " + $nItems[$nSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
        } elseif ($k.Key -eq [ConsoleKey]::Enter -or $k.Key -eq [ConsoleKey]::Escape) {
            if ($k.Key -ne [ConsoleKey]::Escape -and $nSel -eq 0) { Clear-Host; .\Main.ps1 }
            else { [Console]::SetCursorPosition(0, $nTop + 3); Write-Host "  Exiting..." -ForegroundColor DarkYellow }
            return
        }
    }
}

Write-Host ""

# -- Move / Backup files ------------------------------------------------------
$isBackup = ($mSel -eq 2)
$opLabel  = if ($isBackup) { "Copied" } else { "Moved" }
$opVerb   = if ($isBackup) { "Backup" } else { "Move" }

# -- Conflict check -----------------------------------------------------------
Write-Host "  Checking for existing files at destination..." -ForegroundColor DarkCyan
$existsAtDest = @{}
foreach ($f in $rawFiles) {
    $rel    = $f.FullName.Substring($source.Length)
    $target = Join-Path $archiveDest $rel
    if (Test-Path $target) { $existsAtDest[$f.FullName] = $true }
}
$conflicts = $existsAtDest.Count

$overwrite = $false
if ($conflicts -gt 0) {
    Write-Host "  $conflicts file(s) already exist at destination." -ForegroundColor Yellow
    Write-Host ""
    $cfItems = @("Skip existing files", "Overwrite existing files")
    $cfSel   = 0
    $cfTop   = [Console]::CursorTop
    for ($ci = 0; $ci -lt $cfItems.Count; $ci++) {
        [Console]::SetCursorPosition(0, $cfTop + $ci)
        $cfText = ("    " + $cfItems[$ci]).PadRight($w + 4)
        if ($ci -eq $cfSel) { Write-Host $cfText -ForegroundColor Black -BackgroundColor Cyan -NoNewline }
        else                 { Write-Host $cfText -ForegroundColor DarkCyan -NoNewline }
    }
    [Console]::SetCursorPosition(0, $cfTop + $cfItems.Count)
    :cfLoop while ($true) {
        $cfk = [Console]::ReadKey($true)
        if ($cfk.Key -eq [ConsoleKey]::UpArrow -or $cfk.Key -eq [ConsoleKey]::DownArrow) {
            $cfp   = $cfSel; $cfSel = 1 - $cfSel
            [Console]::SetCursorPosition(0, $cfTop + $cfp);   Write-Host ("    " + $cfItems[$cfp]).PadRight($w + 4) -ForegroundColor DarkCyan -NoNewline
            [Console]::SetCursorPosition(0, $cfTop + $cfSel); Write-Host ("    " + $cfItems[$cfSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
        } elseif ($cfk.Key -eq [ConsoleKey]::Enter) {
            break cfLoop
        }
    }
    [Console]::SetCursorPosition(0, $cfTop + $cfItems.Count)
    Write-Host ""
    $overwrite = ($cfSel -eq 1)
}

$done      = 0
$skipped   = 0
$failed    = 0
$errors    = @()
$movedFiles = [System.Collections.Generic.List[object]]::new()

Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  $opVerb - $($rawFiles.Count) file(s)  $totalSizeGB GB" -ForegroundColor Cyan
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host ""

foreach ($f in $rawFiles) {
    $rel       = $f.FullName.Substring($source.Length)
    $target    = Join-Path $archiveDest $rel
    $targetDir = Split-Path $target -Parent

    if ($existsAtDest.ContainsKey($f.FullName) -and -not $overwrite) {
        $skipped++
        Write-Host "  Skipped: $($f.Name)" -ForegroundColor DarkGray
        continue
    }

    try {
        [System.IO.Directory]::CreateDirectory($targetDir) | Out-Null
        if ($isBackup) {
            Copy-Item -Path $f.FullName -Destination $target -Force -ErrorAction Stop
        } else {
            Move-Item -Path $f.FullName -Destination $target -Force -ErrorAction Stop
        }
        $done++
        $movedFiles.Add($f)
        Write-Host "  ${opLabel}: $($f.Name)" -ForegroundColor DarkGray
    } catch {
        $failed++
        $errors += "  FAILED: $($f.FullName) -> $_"
        Write-Host "  FAILED: $($f.Name)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  Done.  ${opLabel}: $done   Skipped: $skipped   Failed: $failed" -ForegroundColor $(if ($failed -eq 0) { "Green" } else { "Yellow" })
if ($errors.Count -gt 0) {
    Write-Host ""
    $errors | ForEach-Object { Write-Host $_ -ForegroundColor Red }
}
Write-Host "  $rule" -ForegroundColor DarkCyan

# -- Update project_info.json -------------------------------------------------
if ($done -gt 0) {
    Write-Host ""
    Write-Host "  Updating project_info.json..." -ForegroundColor DarkCyan

    # Collect unique project folders from actually moved/copied files only
    $affectedProjects = @{}
    foreach ($f in $movedFiles) {
        $dir = $f.DirectoryName
        while ($dir -and $dir.Length -gt $root.Length) {
            $jsonPath = Join-Path $dir "project_info.json"
            if (Test-Path $jsonPath) {
                $affectedProjects[$jsonPath] = $true
                break
            }
            $dir = Split-Path $dir -Parent
        }
    }

    $archiveMode = if ($isBackup) { "Backup" } else { "Move" }
    $archiveDate = (Get-Date -Format "yyyy-MM-dd")
    $tagged = 0

    foreach ($jsonPath in $affectedProjects.Keys) {
        try {
            $info = Get-Content $jsonPath -Raw | ConvertFrom-Json
            $info | Add-Member -NotePropertyName "RawArchived"    -NotePropertyValue $true          -Force
            $info | Add-Member -NotePropertyName "RawArchiveMode" -NotePropertyValue $archiveMode   -Force
            $info | Add-Member -NotePropertyName "RawArchiveDest" -NotePropertyValue $archiveDest   -Force
            $info | Add-Member -NotePropertyName "RawArchiveDate" -NotePropertyValue $archiveDate   -Force
            $info | ConvertTo-Json -Depth 5 | Out-File $jsonPath -Encoding UTF8
            $tagged++
            Write-Host "  Tagged: $jsonPath" -ForegroundColor DarkGray
        } catch {
            Write-Host "  Could not update: $jsonPath" -ForegroundColor Yellow
        }
    }

    Write-Host "  $tagged project(s) updated." -ForegroundColor Cyan
}

# -- Navigation ---------------------------------------------------------------
$nItems = @("Back to main menu", "Exit")
$nSel   = 0
Write-Host ""
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
