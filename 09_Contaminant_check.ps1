$w      = 55
$border = "=" * $w
$rule   = "-" * $w

. (Join-Path $PSScriptRoot "lib\Menu.ps1")

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "   [9]  Contaminant check             (mzsniffer)" -ForegroundColor Cyan
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host ""

# -- Confirm ------------------------------------------------------------------
$r = Show-Menu -Items @("Run", "Back to main menu") -AllowEscape
if ($r.Action -ne "select" -or $r.Index -eq 1) { Return-ToMain; return }
Write-Host ""

$first_path    = Get-Location
$mzsnifferPath = Join-Path $PSScriptRoot "mzsniffer\mzsniffer.exe"

if (-not (Test-Path $mzsnifferPath)) {
    Write-Host "  mzsniffer.exe not found at: $mzsnifferPath" -ForegroundColor Red
    Write-Host "  Place mzsniffer.exe under the mzsniffer\ subfolder next to this script." -ForegroundColor Yellow
    Show-NavExit; return
}

. (Join-Path $PSScriptRoot "lib\Pickers.ps1")
$path = Read-FolderPath "Select the folder containing the .mzML files"
if ($path -eq "") { Return-ToMain; return }
Set-location -Path $path

$files           = Get-ChildItem "*.mzML" | Sort-Object LastWriteTime
$logFilePath     = ".\contaminant_check.txt"
$summaryFilePath = ".\contaminant_summary.csv"

if ($files.Count -eq 0) {
    Write-Host "  No .mzML files found in $path" -ForegroundColor Yellow
    Set-location $first_path
    Write-Host ""
    Write-Host "  $rule" -ForegroundColor DarkCyan
    Show-NavExit
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
Write-Host ""
Write-Host "  $rule" -ForegroundColor DarkCyan
Show-NavExit
return
