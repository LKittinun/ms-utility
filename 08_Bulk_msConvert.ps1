$w      = 55
$border = "=" * $w
$rule   = "-" * $w

. (Join-Path $PSScriptRoot "lib\Menu.ps1")

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "   [8]  Bulk convert .raw to mzML     (msConvert)" -ForegroundColor Cyan
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host ""

# -- Confirm ------------------------------------------------------------------
$r = Show-Menu -Items @("Run", "Back to main menu") -AllowEscape
if ($r.Action -ne "select" -or $r.Index -eq 1) { Return-ToMain; return }
Write-Host ""

if (-not (Get-Command msconvert -ErrorAction SilentlyContinue)) {
    Write-Host "  [ERROR] msconvert not found on PATH. Install ProteoWizard and ensure it is on PATH." -ForegroundColor Red
    Write-Host ""
    pause
    return
}

$first_path = Get-Location
. (Join-Path $PSScriptRoot "lib\Pickers.ps1")
$path = Read-FolderPath "Select the folder containing the .raw files"
if ($path -eq "") { Return-ToMain; return }
Set-location -Path $path

$demux = Read-Host "Demultiplex? y = yes, otherwise = no"
$outputDir = [System.IO.Path]::Combine($path, "mzML_files")
Write-Host $outputDir

if (-Not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir
}

Get-ChildItem -Recurse -Filter *.raw | ForEach-Object {
    $outputFile = Join-Path $outputDir "$($_.BaseName).mzML"
    if (Test-Path $outputFile) {
        Write-Host "File $outputFile already exists. Skipping conversion."
    }
    elseif ($demux -eq "y") {
        msconvert $_.FullName --mzML --outdir $outputDir  --zlib --filter "peakPicking vendor msLevel=1-" --filter "zeroSamples removeExtra 1-" --filter "demultiplex optimization=overlap_only" -v
    }
    else {
        msconvert $_.FullName --mzML --outdir $outputDir  --zlib --filter "peakPicking vendor msLevel=1-" --filter "zeroSamples removeExtra 1-" -v
    }
}

Write-Host "Conversion process complete."
Set-Location -Path $first_path

# ── Navigation ────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  $rule" -ForegroundColor DarkCyan
Show-NavExit
return
