$w      = 55
$border = "=" * $w
$rule   = "-" * $w

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "   [8]  Contaminant check             (mzsniffer)" -ForegroundColor Cyan
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

$files           = Get-ChildItem "*.mzML" | Sort-Object LastWriteTime
$logFilePath     = ".\contaminant_check.txt"
$summaryFilePath = ".\contaminant_summary.csv"
$mzsnifferPath   = Join-Path -Path $first_path -ChildPath "mzsniffer\mzsniffer.exe"

if ($files.Count -eq 0) {
    Write-Host "  No .mzML files found in $path" -ForegroundColor Yellow
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

if (Test-Path $logFilePath) {
    Write-Warning "Replacing existing $logFilePath"
    Remove-Item $logFilePath -Force
}

Write-Host "  Found $($files.Count) file(s). Starting mzsniffer ..."
Write-Host ""

$allPolymers  = [System.Collections.Generic.List[string]]::new()
$fileDataList = @()

foreach ($file in $files) {
    $d = [datetime](Get-ItemProperty -Path $file -Name LastWriteTime).lastwritetime

    Write-Host "  $rule" -ForegroundColor DarkCyan
    Write-Host "  $($file.Name)" -ForegroundColor White
    Write-Host "  Last written : $d"
    Write-Host ""

    $rawOutput = & $mzsnifferPath $file.FullName 2>&1

    "=== $($file.Name)  |  $d ===" | Out-File -FilePath $logFilePath -Append
    $rawOutput                      | Out-File -FilePath $logFilePath -Append
    ""                              | Out-File -FilePath $logFilePath -Append

    $polymerData = @{}
    $dataLines = $rawOutput | Where-Object {
        $_ -match '^\[INFO \] (.+?)\s{2,}([\d]+\.[\d]+)\s*$'
    }

    if ($dataLines.Count -eq 0) {
        Write-Host "    (no data rows parsed)" -ForegroundColor Yellow
    } else {
        Write-Host "    $('Polymer'.PadRight(35)) %TIC"
        Write-Host "    $('-' * 35) ------"
        foreach ($line in $dataLines) {
            $null    = $line -match '^\[INFO \] (.+?)\s{2,}([\d]+\.[\d]+)\s*$'
            $polymer = $matches[1].Trim()
            $tic     = [double]$matches[2]
            $color   = if ($tic -ge 1.0) { "Red" } elseif ($tic -ge 0.1) { "Yellow" } else { "Gray" }
            Write-Host "    $($polymer.PadRight(35)) $tic" -ForegroundColor $color
            $polymerData[$polymer] = $tic
            if (-not $allPolymers.Contains($polymer)) { $allPolymers.Add($polymer) }
        }
    }
    Write-Host ""
    $fileDataList += @{ File = $file.Name; Date = $d; Data = $polymerData }
}

$csvRows = foreach ($fd in $fileDataList) {
    $row = [ordered]@{ File = $fd.File; LastWritten = $fd.Date }
    foreach ($p in $allPolymers) { $row[$p] = if ($fd.Data.ContainsKey($p)) { $fd.Data[$p] } else { "" } }
    [PSCustomObject]$row
}
$csvRows | Export-Csv -Path $summaryFilePath -NoTypeInformation

Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  $($files.Count) file(s) processed." -ForegroundColor Cyan
Write-Host "  Log     : $logFilePath"     -ForegroundColor DarkCyan
Write-Host "  Summary : $summaryFilePath" -ForegroundColor DarkCyan

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
