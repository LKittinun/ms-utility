$w      = 55
$border = "=" * $w
$rule   = "-" * $w
$root   = "Z:\Proteomics\Projects"

Clear-Host
Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "  [11]  Sync from overview CSV" -ForegroundColor Cyan
Write-Host "        Apply edits from overview CSV back to project JSON" -ForegroundColor DarkCyan
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

# -- Locate CSV ---------------------------------------------------------------
$mostRecent = Get-ChildItem -Path $root -Filter "Projects_overview_*.csv" -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1

if ($mostRecent) {
    Write-Host "  Most recent: $($mostRecent.Name)" -ForegroundColor DarkGray
    $csvPath = (Read-Host "  CSV path (blank = use above)").Trim()
    if ($csvPath -eq "") { $csvPath = $mostRecent.FullName }
} else {
    $csvPath = (Read-Host "  CSV path").Trim()
}

function Show-NavExit ($msg) {
    Write-Host ""
    Write-Host "  $msg" -ForegroundColor Yellow
    Write-Host ""
    $nItems = @("Back to main menu", "Exit"); $nSel = 0
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
            return $true
        }
    }
    return $true
}

if (-not (Test-Path $csvPath)) { Show-NavExit "File not found: $csvPath"; return }

# -- Read and validate CSV ----------------------------------------------------
$csvRows = @(Import-Csv -Path $csvPath -ErrorAction SilentlyContinue)
if ($csvRows.Count -eq 0) { Show-NavExit "CSV is empty."; return }

$cols = $csvRows[0].PSObject.Properties.Name
if ($cols -notcontains "ProjectID") { Show-NavExit "CSV missing required column: ProjectID"; return }

# -- Scan project_info.json files ---------------------------------------------
Write-Host ""
Write-Host "  Scanning $root ..." -ForegroundColor DarkCyan

$jsonFiles = @(Get-ChildItem -Path $root -Recurse -Filter "project_info.json" -ErrorAction SilentlyContinue)

$lookup = @{}
foreach ($jf in $jsonFiles) {
    try {
        $info = Get-Content $jf.FullName -Raw | ConvertFrom-Json
        if ($info.ProjectID) { $lookup[$info.ProjectID] = @{ Info = $info; Path = $jf.FullName } }
    } catch { }
}

# -- Detect changes -----------------------------------------------------------
# Fields: CSV column name, JSON property name, display label
$syncFields = @(
    @{ CSV = "PI";                    JSON = "PI";                    Label = "PI"          }
    @{ CSV = "ColumnDescription";     JSON = "ColumnDescription";     Label = "Column desc" }
    @{ CSV = "TrapColumn";            JSON = "TrapColumn";            Label = "Trap column" }
    @{ CSV = "TrapColumnDescription"; JSON = "TrapColumnDescription"; Label = "Trap desc"   }
    @{ CSV = "SampleFolders";         JSON = "SampleFolders";         Label = "Folders"     }
)

$changes = @()
foreach ($row in $csvRows) {
    $id = if ($row.ProjectID) { $row.ProjectID.Trim() } else { "" }
    if ($id -eq "" -or $id -eq "-" -or -not $lookup.ContainsKey($id)) { continue }

    $entry = $lookup[$id]
    $info  = $entry.Info
    $diffs = @()

    foreach ($sf in $syncFields) {
        if ($cols -notcontains $sf.CSV) { continue }

        if ($sf.CSV -eq "SampleFolders") {
            # Array comparison
            $jsonArr = if ($info.SampleFolders) {
                @($info.SampleFolders | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
            } else { @() }
            $csvArr = if ($row.SampleFolders -and $row.SampleFolders -ne "-") {
                @($row.SampleFolders -split ";" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
            } else { @() }
            $jsonKey = ($jsonArr | Sort-Object) -join ";"
            $csvKey  = ($csvArr  | Sort-Object) -join ";"
            if ($jsonKey -ne $csvKey) {
                $diffs += @{
                    Label    = "Folders"
                    JSON     = "SampleFolders"
                    IsArray  = $true
                    Old      = $jsonArr -join "; "
                    New      = $csvArr  -join "; "
                    NewArray = $csvArr
                }
            }
        } else {
            $jsonVal = if ($info.$($sf.JSON)) { [string]$info.$($sf.JSON) } else { "" }
            $csvVal  = if ($row.$($sf.CSV) -and $row.$($sf.CSV) -ne "-") { [string]$row.$($sf.CSV) } else { "" }
            if ($jsonVal -ne $csvVal) {
                $diffs += @{ Label = $sf.Label; JSON = $sf.JSON; Old = $jsonVal; New = $csvVal }
            }
        }
    }

    if ($diffs.Count -gt 0) {
        $changes += @{ ID = $id; Project = $info.Project; Path = $entry.Path; Info = $info; Diffs = $diffs }
    }
}

Write-Host ""

if ($changes.Count -eq 0) {
    Show-NavExit "No changes detected - CSV matches current project data."
    return
}

# -- Preview ------------------------------------------------------------------
Write-Host "  $($changes.Count) project(s) with changes:" -ForegroundColor Cyan
Write-Host ""
foreach ($c in $changes) {
    Write-Host "  $rule" -ForegroundColor DarkCyan
    Write-Host "  $($c.ID)  $($c.Project)" -ForegroundColor White
    foreach ($d in $c.Diffs) {
        $oldStr = if ($d.Old -eq "") { "(empty)" } else { $d.Old }
        $newStr = if ($d.New -eq "") { "(empty)" } else { $d.New }
        Write-Host ("    " + $d.Label.PadRight(14) + ": ") -NoNewline -ForegroundColor DarkCyan
        Write-Host $oldStr -NoNewline -ForegroundColor DarkGray
        Write-Host " -> " -NoNewline -ForegroundColor DarkCyan
        Write-Host $newStr -ForegroundColor White
    }
}
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host ""
Write-Host "  Note: SampleFolders changes update metadata only, not folders on disk." -ForegroundColor DarkGray
Write-Host ""

# -- Apply confirm ------------------------------------------------------------
$aItems = @("Apply changes", "Cancel")
$aSel   = 0
$aTop   = [Console]::CursorTop
[Console]::SetCursorPosition(0, $aTop)
Write-Host ("  > " + $aItems[0]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
[Console]::SetCursorPosition(0, $aTop + 1)
Write-Host ("    " + $aItems[1]).PadRight($w + 4) -ForegroundColor DarkYellow -NoNewline
[Console]::SetCursorPosition(0, $aTop + 2)
while ($true) {
    $k = [Console]::ReadKey($true)
    if ($k.Key -eq [ConsoleKey]::UpArrow -or $k.Key -eq [ConsoleKey]::DownArrow) {
        $p = $aSel; $aSel = 1 - $aSel
        [Console]::SetCursorPosition(0, $aTop + $p)
        Write-Host ("    " + $aItems[$p]).PadRight($w + 4) -ForegroundColor $(if ($p -eq 0) { "Cyan" } else { "DarkYellow" }) -NoNewline
        [Console]::SetCursorPosition(0, $aTop + $aSel)
        Write-Host ("  > " + $aItems[$aSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
        [Console]::SetCursorPosition(0, $aTop + 2)
    } elseif ($k.Key -eq [ConsoleKey]::Enter -or $k.Key -eq [ConsoleKey]::Escape) {
        [Console]::SetCursorPosition(0, $aTop + 2)
        if ($k.Key -eq [ConsoleKey]::Escape -or $aSel -eq 1) {
            Show-NavExit "Cancelled - no changes applied."
            return
        }
        break
    }
}
Write-Host ""

# -- Apply --------------------------------------------------------------------
$updated = 0
foreach ($c in $changes) {
    $info = $c.Info

    foreach ($d in $c.Diffs) {
        $newVal = if ($d.New -eq "") { $null } else { $d.New }
        if ($d.IsArray) {
            $info | Add-Member -NotePropertyName "SampleFolders" -NotePropertyValue @($d.NewArray) -Force
        } else {
            $info | Add-Member -NotePropertyName $d.JSON -NotePropertyValue $newVal -Force
        }
    }

    $info | ConvertTo-Json | Out-File $c.Path -Encoding UTF8

    # Update column_log.csv if it exists
    $logFile = Join-Path (Split-Path (Split-Path $c.Path)) "column_log.csv"
    if (Test-Path $logFile) {
        $logRows = @(Import-Csv $logFile)
        $matchIdx = -1
        for ($ri = 0; $ri -lt $logRows.Count; $ri++) {
            if ($logRows[$ri].ProjectID -eq $c.ID) { $matchIdx = $ri; break }
        }
        if ($matchIdx -ge 0) {
            foreach ($d in $c.Diffs) {
                if ($d.IsArray) {
                    $logRows[$matchIdx].SampleFolders = $d.NewArray -join ";"
                } else {
                    $csvCol = ($syncFields | Where-Object { $_.JSON -eq $d.JSON }).CSV
                    if ($csvCol -and ($logRows[$matchIdx].PSObject.Properties.Name -contains $csvCol)) {
                        $logRows[$matchIdx].$csvCol = $d.New
                    }
                }
            }
            $logRows | Export-Csv $logFile -NoTypeInformation
        }
    }

    Write-Host "  Updated: $($c.ID)  $($c.Project)" -ForegroundColor Green
    $updated++
}

Write-Host ""
Write-Host "  $updated project(s) updated." -ForegroundColor Cyan
Write-Host ""

# -- Navigation ---------------------------------------------------------------
$nItems = @("Back to main menu", "Exit")
$nSel   = 0
$nTop   = [Console]::CursorTop
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
