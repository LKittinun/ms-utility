$w      = 55
$border = "=" * $w
$rule   = "-" * $w

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "   [4]  Column usage report" -ForegroundColor Cyan
Write-Host "        All .raw files must be within column parent dir" -ForegroundColor DarkCyan
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

$first_path = Get-location
$path = Read-Host "Insert directory, leave blank for a current location"
if ($path -eq "") { $path = (Get-Location).Path }
Set-location -Path $path

$currentDate        = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFilePath     = [System.IO.Path]::Combine($path, "Column_usage_history")
$outputCsvFilePath  = "$outputFilePath\Column_info_$currentDate.csv"
$outputTextFilePath = "$outputFilePath\Column_info_$currentDate.txt"

if (-Not (Test-Path -Path $outputFilePath)) {
    New-Item -Path $outputFilePath -ItemType Directory > $null
}

$files = Get-ChildItem *.raw -Path $path -File -Recurse | Sort-Object CreationTime

if ($files.Count -eq 0) {
    Write-Host "  No .raw files found." -ForegroundColor Yellow
    Set-location $first_path
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

$runNo = 0
$fileInfoList = foreach ($file in $files) {
    $runNo++
    [PSCustomObject]@{
        RunNo        = $runNo
        Name         = $file.Name
        Subfolder    = $file.Directory.Name
        SizeMB       = [Math]::Round($file.Length / 1MB, 2)
        CreationTime = $file.CreationTime
        FullPath     = $file.FullName
    }
}

$filesBlank   = $fileInfoList | Where-Object { $_.Subfolder -match "(?i)^blank$" }
$filesPRTC    = $fileInfoList | Where-Object { $_.Subfolder -match "(?i)^prtc$" }
$filesSamples = $fileInfoList | Where-Object { $_.Subfolder -notmatch "(?i)^(blank|prtc)$" }

$firstFile        = $fileInfoList  | Select-Object -First 1
$lastFile         = $fileInfoList  | Select-Object -Last 1
$firstFileSample  = $filesSamples  | Select-Object -First 1
$lastFileSample   = $filesSamples  | Select-Object -Last 1
$spanDays  = [Math]::Round(($lastFile.CreationTime - $firstFile.CreationTime).TotalDays, 1)
$avgPerDay = if ($spanDays -gt 0) { [Math]::Round($filesSamples.Count / $spanDays, 1) } else { "N/A" }
$totalGB   = [Math]::Round(($fileInfoList | Measure-Object SizeMB -Sum).Sum / 1024, 2)
$minSizeMB = ($fileInfoList | Measure-Object SizeMB -Minimum).Minimum
$maxSizeMB = ($fileInfoList | Measure-Object SizeMB -Maximum).Maximum
$medSizeMB = ($fileInfoList.SizeMB | Sort-Object)[[Math]::Floor($fileInfoList.Count / 2)]

$subfolderCounts = $fileInfoList |
    Where-Object { $_.Subfolder -ne "Column_usage_history" } |
    Group-Object Subfolder |
    Sort-Object Count -Descending |
    Select-Object Name, Count

$fileInfoList | Export-Csv -Path $outputCsvFilePath -NoTypeInformation

$summary = @(
    "Generated       : $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    "Directory       : $path"
    ""
    "--- Injection count ---"
    "Total runs      : $($fileInfoList.Count)"
    "Blank           : $($filesBlank.Count)"
    "PRTC            : $($filesPRTC.Count)"
    "Samples         : $($filesSamples.Count)"
    ""
    "--- Timeline ---"
    "First run       : $($firstFile.CreationTime.ToString('yyyy-MM-dd HH:mm'))  $($firstFile.Name)"
    "Last run        : $($lastFile.CreationTime.ToString('yyyy-MM-dd HH:mm'))  $($lastFile.Name)"
    "Span            : $spanDays days"
    "Avg / day       : $avgPerDay runs (samples only)"
    ""
    "--- Data volume ---"
    "Total size      : $totalGB GB"
    "File size min   : $minSizeMB MB"
    "File size median: $medSizeMB MB"
    "File size max   : $maxSizeMB MB"
    ""
    "--- Runs per subfolder ---"
)
if (($filesBlank.Count -gt 0) -or ($filesPRTC.Count -gt 0)) {
    $summary += "First sample    : $($firstFileSample.CreationTime.ToString('yyyy-MM-dd HH:mm'))  $($firstFileSample.Name)"
    $summary += "Last sample     : $($lastFileSample.CreationTime.ToString('yyyy-MM-dd HH:mm'))  $($lastFileSample.Name)"
}
foreach ($sf in $subfolderCounts) { $summary += "$($sf.Name.PadRight(20)): $($sf.Count)" }
$summary | Out-File -FilePath $outputTextFilePath -Encoding UTF8

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "  Results" -ForegroundColor Cyan
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  Total runs   : $($fileInfoList.Count)"
Write-Host "  Blank        : $($filesBlank.Count)"
Write-Host "  PRTC         : $($filesPRTC.Count)"
Write-Host "  Samples      : $($filesSamples.Count)"
Write-Host "  First run    : $($firstFile.CreationTime.ToString('yyyy-MM-dd'))  $($firstFile.Name)"
Write-Host "  Last run     : $($lastFile.CreationTime.ToString('yyyy-MM-dd'))  $($lastFile.Name)"
if (($filesBlank.Count -gt 0) -or ($filesPRTC.Count -gt 0)) {
    Write-Host "  First sample : $($firstFileSample.CreationTime.ToString('yyyy-MM-dd'))  $($firstFileSample.Name)"
    Write-Host "  Last sample  : $($lastFileSample.CreationTime.ToString('yyyy-MM-dd'))  $($lastFileSample.Name)"
}
Write-Host "  Span         : $spanDays days   avg $avgPerDay runs/day (samples)"
Write-Host "  Total data   : $totalGB GB"
Write-Host "  File size    : min $minSizeMB  med $medSizeMB  max $maxSizeMB MB"
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  Runs per subfolder" -ForegroundColor Cyan
Write-Host "  $rule" -ForegroundColor DarkCyan
foreach ($sf in $subfolderCounts) { Write-Host "  $($sf.Name.PadRight(30)) $($sf.Count) runs" }
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  CSV : $outputCsvFilePath" -ForegroundColor DarkCyan
Write-Host "  TXT : $outputTextFilePath" -ForegroundColor DarkCyan
Write-Host "  $border" -ForegroundColor DarkCyan

Set-location $first_path

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
