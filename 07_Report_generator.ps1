$w      = 55
$border = "=" * $w
$rule   = "-" * $w

. (Join-Path $PSScriptRoot "lib\Menu.ps1")

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "   [7]  Analysis report               (Excel)" -ForegroundColor Cyan
Write-Host "        Analysis_Report.xlsx  (5 sheets)" -ForegroundColor DarkCyan
Write-Host "        Project Overview  |  Raw Files  |  Run Statistics" -ForegroundColor DarkCyan
Write-Host "        Summary Statistics  |  Protein Groups (pg_matrix)" -ForegroundColor DarkCyan
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host ""

# -- Confirm ------------------------------------------------------------------
$r = Show-Menu -Items @("Run", "Back to main menu") -AllowEscape
if ($r.Action -ne "select" -or $r.Index -eq 1) { Return-ToMain; return }
Write-Host ""

# -- Folder selection ---------------------------------------------------------
. (Join-Path $PSScriptRoot "lib\Pickers.ps1")
$path = Read-FolderPath "Select the project folder"
if ($path -eq "") { Return-ToMain; return }
Write-Host ""

# Optional extra folder holding raw files (e.g. archived to external HDD)
$path2  = ""
Write-Host "  Raw files in another folder too? (e.g. archived to external HDD)" -ForegroundColor Cyan
$r    = Show-Menu -Items @("No", "Yes - select folder")
$xSel = $r.Index
Write-Host ""
if ($xSel -eq 1) {
    $path2 = Read-FolderPath "Select the extra folder containing raw files" -NoStart
    if ($path2 -eq "") { Write-Host "  No extra folder selected" -ForegroundColor DarkGray }
}
Write-Host ""

if (-not (Test-Path -LiteralPath $path -PathType Container)) {
    Write-Host "  ERROR: Project folder not found: $path" -ForegroundColor Red
    $path = $null
}
if ($path -and $path2 -ne "") {
    if (-not (Test-Path -LiteralPath $path2 -PathType Container)) {
        Write-Host "  WARNING: Extra raw files folder not found, ignoring: $path2" -ForegroundColor Yellow
        $path2 = ""
    } elseif ((Resolve-Path -LiteralPath $path2).Path.TrimEnd("\") -eq (Resolve-Path -LiteralPath $path).Path.TrimEnd("\")) {
        Write-Host "  WARNING: Extra raw files folder is the same as the project folder, ignoring" -ForegroundColor Yellow
        $path2 = ""
    } else {
        Write-Host "  Extra raw files folder : $path2" -ForegroundColor DarkCyan
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
                    Write-Host "  No '$($sf.Name)' subfolder in extra folder - extra raw files skipped for this sample" -ForegroundColor Yellow
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
Write-Host ""
Write-Host "  $rule" -ForegroundColor DarkCyan
Show-NavExit
return
