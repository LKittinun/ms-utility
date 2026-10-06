$w      = 55
$border = "=" * $w
$rule   = "-" * $w

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "   [7]  Analysis report               (Excel)" -ForegroundColor Cyan
Write-Host "        Analysis_Report.xlsx  (5 sheets)" -ForegroundColor DarkCyan
Write-Host "        Project Overview  |  Raw Files  |  Run Statistics" -ForegroundColor DarkCyan
Write-Host "        Summary Statistics  |  Protein Groups (pg_matrix)" -ForegroundColor DarkCyan
Write-Host "  $border" -ForegroundColor DarkCyan
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

$first_path = Get-Location
$raw = Read-Host "Insert project directory, leave blank for current location (use ; to chain two folders)"
if ($raw -eq "") { $raw = $first_path.Path }
$parts = @($raw -split ";" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
if ($parts.Count -eq 0) { $parts = @($first_path.Path) }
$path  = $parts[0]
$path2 = if ($parts.Count -ge 2) { $parts[1] } else { "" }
if (-not (Test-Path -LiteralPath $path -PathType Container)) {
    Write-Host "  ERROR: Project folder not found: $path" -ForegroundColor Red
    $path = $null
}
if ($path -and $path2 -ne "") {
    if (-not (Test-Path -LiteralPath $path2 -PathType Container)) {
        Write-Host "  WARNING: Chained folder not found, ignoring: $path2" -ForegroundColor Yellow
        $path2 = ""
    } elseif ((Resolve-Path -LiteralPath $path2).Path.TrimEnd("\") -eq (Resolve-Path -LiteralPath $path).Path.TrimEnd("\")) {
        Write-Host "  WARNING: Chained folder is the same as the project folder, ignoring" -ForegroundColor Yellow
        $path2 = ""
    } else {
        Write-Host "  Chained folder : $path2" -ForegroundColor DarkCyan
    }
    Write-Host ""
}

# Detect subfolders that contain a Result\ subfolder (sample type folders)
$subfolders = @()
if ($path) {
    $subfolders = @(Get-ChildItem -LiteralPath $path -Directory |
        Where-Object { Test-Path (Join-Path $_.FullName "Result") })
}

if (-not $path) {
    # Invalid project folder - skip processing, fall through to navigation
} elseif ($subfolders.Count -eq 0) {
    # No subfolders with Result\ found - treat $path itself as the project folder
    $subfolders = @([PSCustomObject]@{ FullName = $path; Name = Split-Path $path -Leaf })
    Write-Host "  No subfolders with Result\ found - running on: $path" -ForegroundColor Yellow
} else {
    Write-Host "  Found $($subfolders.Count) subfolder(s):" -ForegroundColor Cyan
    $subfolders | ForEach-Object { Write-Host "    $($_.Name)" -ForegroundColor White }
    Write-Host ""
}

# Locate Rscript.exe
$rscript = $null
try { $null = & Rscript --version 2>&1; $rscript = "Rscript" } catch {}
if (-not $rscript) {
    $rBase = "C:\Program Files\R"
    if (Test-Path $rBase) {
        $candidates = Get-ChildItem -Path $rBase -Directory |
                      Sort-Object Name -Descending |
                      ForEach-Object { Join-Path $_.FullName "bin\Rscript.exe" } |
                      Where-Object { Test-Path $_ }
        if ($candidates) { $rscript = $candidates[0] }
    }
}

if (-not $path) {
    # Nothing to process
} elseif (-not $rscript) {
    Write-Host "ERROR: Rscript.exe not found. Install R from https://cran.r-project.org/" -ForegroundColor Red
} else {
    Write-Host "Using R: $rscript" -ForegroundColor Cyan
    foreach ($sf in $subfolders) {
        $sfPath = $sf.FullName.Replace("\", "/")
        Write-Host ""
        Write-Host "  $rule" -ForegroundColor DarkCyan
        Write-Host "  Processing: $($sf.Name)" -ForegroundColor Cyan
        Write-Host "  $rule" -ForegroundColor DarkCyan
        try {
            $rArgs = @(".\R\generate_report.R", $sfPath, "Result", $sfPath)
            # Resolve the chained raw-file folder for this sample folder:
            # prefer a same-named subfolder; fall back to the chained root only
            # when there is a single sample folder (otherwise every sample would
            # pick up the root's files)
            $sfChained = ""
            if ($path2 -ne "") {
                $sub = Join-Path $path2 $sf.Name
                if (Test-Path -LiteralPath $sub -PathType Container) {
                    $sfChained = $sub
                } elseif ($subfolders.Count -le 1) {
                    $sfChained = $path2
                } else {
                    Write-Host "  No '$($sf.Name)' subfolder in chained folder - chained raw files skipped for this sample" -ForegroundColor Yellow
                }
            }
            if ($sfChained -ne "") { $rArgs += $sfChained.Replace("\", "/") }
            & $rscript @rArgs 2>&1 | ForEach-Object { "$_" }
        } catch {
            Write-Host "  Unexpected error during R execution: $_" -ForegroundColor Red
        }
        if ($LASTEXITCODE -ne 0) {
            Write-Host "  R script exited with code $LASTEXITCODE -- check output above for details." -ForegroundColor Red
        }
    }
}

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
