$w      = 55
$border = "=" * $w
$rule   = "-" * $w
. (Join-Path $PSScriptRoot "lib\Menu.ps1")

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "   [4]  Column usage report" -ForegroundColor Cyan
Write-Host "        All .raw files must be within column parent dir" -ForegroundColor DarkCyan
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host ""

# -- Confirm ------------------------------------------------------------------
$r = Show-Menu -Items @("Run", "Back to main menu") -AllowEscape
if ($r.Action -ne "select" -or $r.Index -eq 1) { Return-ToMain; return }
Write-Host ""

$first_path = Get-location
. (Join-Path $PSScriptRoot "lib\Pickers.ps1")
$path = Read-FolderPath "Select the folder to scan for .raw files"
if ($path -eq "") { Return-ToMain; return }
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
    Show-NavExit
    return
}

# Build a map from each sample folder full path -> project info, using SampleFolders from json
$sampleDirMap = @{}
Get-ChildItem -Path $path -Recurse -Filter "project_info.json" -File | ForEach-Object {
    $info    = Get-Content $_.FullName -Raw | ConvertFrom-Json
    $projDir = $_.DirectoryName
    $folders = @($info.SampleFolders)
    if ($folders.Count -eq 0) {
        $sampleDirMap[$projDir] = $info
    } else {
        foreach ($sf in $folders) {
            $sampleDirMap[(Join-Path $projDir $sf)] = $info
        }
    }
}

$runNo = 0
$fileInfoList = foreach ($file in $files) {
    $runNo++
    $projInfo   = $sampleDirMap[$file.DirectoryName]
    $trapCol    = if ($projInfo -and $projInfo.TrapColumn)            { $projInfo.TrapColumn            } else { "" }
    $trapDesc   = if ($projInfo -and $projInfo.TrapColumnDescription) { $projInfo.TrapColumnDescription } else { "" }
    $grandParent = $file.Directory.Parent
    $grandName   = $grandParent.Name -replace '^\d{4}-\d{2}-\d{2}_', ''
    $leafName    = $file.Directory.Name
    $subLabel    = if ($grandParent.FullName -eq $path) {
                       $leafName -replace '^\d{4}-\d{2}-\d{2}_', ''
                   } else {
                       "$grandName/$leafName"
                   }
    [PSCustomObject]@{
        RunNo                 = $runNo
        Name                 = $file.Name
        Subfolder            = $leafName
        SubfolderLabel       = $subLabel
        SizeMB               = [Math]::Round($file.Length / 1MB, 2)
        CreationTime         = $file.CreationTime
        TrapColumn           = $trapCol
        TrapColumnDescription = $trapDesc
        FullPath             = $file.FullName
    }
}

$filesBlank    = $fileInfoList | Where-Object { $_.Subfolder -match "(?i)^blank$" }
$filesPRTC     = $fileInfoList | Where-Object { $_.Subfolder -match "(?i)^prtc$" }
$filesSamples  = $fileInfoList | Where-Object { $_.Subfolder -notmatch "(?i)^(blank|prtc)$" }
$filesNonBlank = $fileInfoList | Where-Object { $_.Subfolder -notmatch "(?i)^blank$" }

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
    Group-Object SubfolderLabel |
    Sort-Object Count -Descending |
    Select-Object Name, Count

$trapColGroups = $fileInfoList |
    Where-Object { $_.Subfolder -notmatch "(?i)^blank$" -and $_.TrapColumn -ne "" } |
    Group-Object TrapColumn |
    Sort-Object Count -Descending |
    Select-Object Name, Count, @{N='Description'; E={($_.Group | Select-Object -First 1).TrapColumnDescription}}
$trapUnclassified = ($fileInfoList | Where-Object { $_.Subfolder -notmatch "(?i)^blank$" -and $_.TrapColumn -eq "" }).Count

$fileInfoList | Export-Csv -Path $outputCsvFilePath -NoTypeInformation

$summary = @(
    "Generated       : $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    "Directory       : $path"
    ""
    "--- Injection count ---"
    "Total runs      : $($filesNonBlank.Count)  (blank excluded)"
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
if ((($filesBlank.Count -gt 0) -or ($filesPRTC.Count -gt 0)) -and ($filesSamples.Count -gt 0)) {
    $summary += "First sample    : $($firstFileSample.CreationTime.ToString('yyyy-MM-dd HH:mm'))  $($firstFileSample.Name)"
    $summary += "Last sample     : $($lastFileSample.CreationTime.ToString('yyyy-MM-dd HH:mm'))  $($lastFileSample.Name)"
}
foreach ($sf in $subfolderCounts) { $summary += "$($sf.Name.PadRight(20)): $($sf.Count)" }
$summary += ""
$summary += "--- Trap column usage ---"
if ($trapColGroups.Count -eq 0 -and $trapUnclassified -eq $fileInfoList.Count) {
    $summary += "(no project_info.json found - trap column unknown)"
} else {
    foreach ($tc in $trapColGroups) {
        $label = if ($tc.Description) { "$($tc.Name)  ($($tc.Description))" } else { $tc.Name }
        $summary += "$($label.PadRight(48)): $($tc.Count) runs"
    }
    if ($trapUnclassified -gt 0) { $summary += "(unclassified)$(' ' * 34): $trapUnclassified runs" }
}
$summary | Out-File -FilePath $outputTextFilePath -Encoding UTF8

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "  Results" -ForegroundColor Cyan
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  Total runs   : $($filesNonBlank.Count)  (blank excluded)"
Write-Host "  Blank        : $($filesBlank.Count)"
Write-Host "  PRTC         : $($filesPRTC.Count)"
Write-Host "  Samples      : $($filesSamples.Count)"
Write-Host "  First run    : $($firstFile.CreationTime.ToString('yyyy-MM-dd'))  $($firstFile.Name)"
Write-Host "  Last run     : $($lastFile.CreationTime.ToString('yyyy-MM-dd'))  $($lastFile.Name)"
if ((($filesBlank.Count -gt 0) -or ($filesPRTC.Count -gt 0)) -and ($filesSamples.Count -gt 0)) {
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
Write-Host "  Trap column usage" -ForegroundColor Cyan
Write-Host "  $rule" -ForegroundColor DarkCyan
if ($trapColGroups.Count -eq 0 -and $trapUnclassified -eq $fileInfoList.Count) {
    Write-Host "  (no project_info.json found - trap column unknown)" -ForegroundColor DarkGray
} else {
    foreach ($tc in $trapColGroups) {
        $label = if ($tc.Description) { "$($tc.Name)  ($($tc.Description))" } else { $tc.Name }
        Write-Host "  $($label.PadRight(40)) $($tc.Count) runs"
    }
    if ($trapUnclassified -gt 0) { Write-Host "  (unclassified)$(' ' * 26) $trapUnclassified runs" -ForegroundColor DarkGray }
}
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  CSV : $outputCsvFilePath" -ForegroundColor DarkCyan
Write-Host "  TXT : $outputTextFilePath" -ForegroundColor DarkCyan
Write-Host "  $border" -ForegroundColor DarkCyan

Set-location $first_path

# ── Navigation ────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  $rule" -ForegroundColor DarkCyan
Show-NavExit
return
