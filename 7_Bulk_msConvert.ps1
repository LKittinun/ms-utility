$w      = 55
$border = "=" * $w
$rule   = "-" * $w

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "   [7]  Bulk convert .raw to mzML     (msConvert)" -ForegroundColor Cyan
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host ""

$first_path = Get-Location
$path = Read-Host "Set the path to the directory containing the .raw files, leave blank for a current location"
if ($path -eq "") { $path = (Get-Location).Path }
Set-location -Path $path

$demux = Read-Host "Demultiplex? y = yes, otherwise = no"
$outputDir = [System.IO.Path]::Combine($path, "mzML_files")
Write-Host $outputDir

if (-Not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir
}

Get-ChildItem -Filter *.raw | ForEach-Object {
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
        [Console]::SetCursorPosition(0, $nTop + $p)
        Write-Host ("    " + $nItems[$p]).PadRight($w + 4) -ForegroundColor $(if ($p -eq 0) { "Cyan" } else { "DarkYellow" }) -NoNewline
        [Console]::SetCursorPosition(0, $nTop + $nSel)
        Write-Host ("  > " + $nItems[$nSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
    } elseif ($k.Key -eq [ConsoleKey]::Enter -or $k.Key -eq [ConsoleKey]::Escape) {
        if ($k.Key -ne [ConsoleKey]::Escape -and $nSel -eq 0) { Clear-Host; .\Main.ps1 }
        else { [Console]::SetCursorPosition(0, $nTop + 3); Write-Host "  Exiting..." -ForegroundColor DarkYellow }
        return
    }
}
