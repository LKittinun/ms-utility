$w      = 55
$border = "=" * $w
$rule   = "-" * $w

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "   [10] Archive raw files              (to external HDD)" -ForegroundColor Cyan
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

# -- Archive destination ------------------------------------------------------
$archiveDest = Read-Host "  Archive destination (e.g. E:\Raw_Archive)"
if ($archiveDest -eq "") {
    Write-Host "  Cancelled." -ForegroundColor DarkYellow
    Write-Host ""
    pause
    Clear-Host; .\Main.ps1; return
}

Write-Host ""
Write-Host "  Scanning for .raw files..." -ForegroundColor DarkCyan

# -- Scan ---------------------------------------------------------------------
$rawFiles = @(Get-ChildItem -Path "$root\*" -Recurse -Filter "*.raw" -File)

if ($rawFiles.Count -eq 0) {
    Write-Host "  No .raw files found under $root" -ForegroundColor Yellow
    Write-Host ""
    pause
    Clear-Host; .\Main.ps1; return
}

$totalSizeBytes = ($rawFiles | Measure-Object -Property Length -Sum).Sum
$totalSizeGB    = [math]::Round($totalSizeBytes / 1GB, 2)

Write-Host "  Found $($rawFiles.Count) .raw file(s)  ($totalSizeGB GB)" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Source : $root" -ForegroundColor DarkGray
Write-Host "  Dest   : $archiveDest" -ForegroundColor DarkGray
Write-Host ""

# -- Mode picker --------------------------------------------------------------
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
        $rel    = $_.FullName.Substring($root.Length)
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
$conflicts = @($rawFiles | Where-Object {
    $rel    = $_.FullName.Substring($root.Length)
    $target = Join-Path $archiveDest $rel
    Test-Path $target
})

$overwrite = $false
if ($conflicts.Count -gt 0) {
    Write-Host "  $($conflicts.Count) file(s) already exist at destination." -ForegroundColor Yellow
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

$done   = 0
$skipped = 0
$failed = 0
$errors = @()

Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  $opVerb - $($rawFiles.Count) file(s)  $totalSizeGB GB" -ForegroundColor Cyan
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host ""

foreach ($f in $rawFiles) {
    $rel       = $f.FullName.Substring($root.Length)
    $target    = Join-Path $archiveDest $rel
    $targetDir = Split-Path $target -Parent

    if ((Test-Path $target) -and -not $overwrite) {
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
