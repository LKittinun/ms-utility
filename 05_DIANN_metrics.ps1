$w      = 55
$border = "=" * $w
$rule   = "-" * $w
. (Join-Path $PSScriptRoot "lib\Menu.ps1")

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "   [5]  DIA-NN metrics                (plots + TSV)" -ForegroundColor Cyan
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host ""

# -- Confirm ------------------------------------------------------------------
$r = Show-Menu -Items @("Run", "Back to main menu") -AllowEscape
if ($r.Action -ne "select" -or $r.Index -eq 1) { Return-ToMain; return }
Write-Host ""

$first_path = Get-Location
. (Join-Path $PSScriptRoot "lib\Pickers.ps1")
$path = Read-FolderPath "Select the project folder"
if ($path -eq "") { Return-ToMain; return }

# Detect subfolders that contain a Result\ subfolder with DIA-NN output
$subfolders = Get-ChildItem -Path $path -Directory |
    Where-Object { Test-Path (Join-Path $_.FullName "Result") } |
    Where-Object {
        (Test-Path (Join-Path $_.FullName "Result\report.parquet")) -or
        (Test-Path (Join-Path $_.FullName "Result\report.tsv"))
    }

if ($subfolders.Count -eq 0) {
    # No subfolders found - check if $path itself contains Result\report.*
    $resultDir = Join-Path $path "Result"
    if ((Test-Path (Join-Path $resultDir "report.parquet")) -or
        (Test-Path (Join-Path $resultDir "report.tsv"))) {
        $subfolders = @([PSCustomObject]@{ FullName = $path; Name = Split-Path $path -Leaf })
        Write-Host "  No subfolders with Result\ found - running on: $path" -ForegroundColor Yellow
    } else {
        Write-Host "  No subfolders with DIA-NN output found in: $path" -ForegroundColor Yellow
        Write-Host "  Expected: <subfolder>\Result\report.parquet (or report.tsv)" -ForegroundColor DarkCyan
    }
}

if ($subfolders.Count -gt 0) {
    Write-Host "  Found $($subfolders.Count) subfolder(s):" -ForegroundColor Cyan
    $subfolders | ForEach-Object { Write-Host "    $($_.Name)" -ForegroundColor White }
    Write-Host ""

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

    if (-not $rscript) {
        Write-Host "ERROR: Rscript.exe not found. Install R from https://cran.r-project.org/" -ForegroundColor Red
    } else {
        Write-Host "Using R: $rscript" -ForegroundColor Cyan
        foreach ($sf in $subfolders) {
            $resultPath = (Join-Path $sf.FullName "Result").Replace("\", "/")
            Write-Host ""
            Write-Host "  $rule" -ForegroundColor DarkCyan
            Write-Host "  Processing: $($sf.Name)" -ForegroundColor Cyan
            Write-Host "  $rule" -ForegroundColor DarkCyan
            try {
                & $rscript ".\R\DIANN_metrics.R" $resultPath 2>&1 | ForEach-Object { "$_" }
            } catch {
                Write-Host "  Unexpected error during R execution: $_" -ForegroundColor Red
            }
            if ($LASTEXITCODE -ne 0) {
                Write-Host "  R script exited with code $LASTEXITCODE -- check output above for details." -ForegroundColor Red
            }
        }
    }
}

# ── Navigation ────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  $rule" -ForegroundColor DarkCyan
Show-NavExit
return
