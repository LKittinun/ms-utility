$w          = 55
$border     = "=" * $w
$rule       = "-" * $w
$root       = "Z:\Proteomics\Projects"
$prohibited = @("blank", "raw_summary", "prtc", "sst", "column_usage_history")

Clear-Host
Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "   [5]  Projects overview" -ForegroundColor Cyan
Write-Host "        Filter and export all projects (Excel)" -ForegroundColor DarkCyan
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

# -- Check ImportExcel --------------------------------------------------------
$hasImportExcel = $null -ne (Get-Module -ListAvailable -Name ImportExcel)
if (-not $hasImportExcel) {
    Write-Host "  ImportExcel module not found." -ForegroundColor Yellow
    Write-Host "  Install with: Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor DarkCyan
    Write-Host "  Falling back to CSV export." -ForegroundColor DarkCyan
    Write-Host ""
}

# -- Scan ---------------------------------------------------------------------
Write-Host "  Scanning $root ..." -ForegroundColor DarkCyan

$jsonFiles = @(Get-ChildItem -Path $root -Recurse -Filter "project_info.json" -ErrorAction SilentlyContinue)
$jsonFiles = @($jsonFiles | Where-Object {
    $prohibited -notcontains ($_.Directory.Name -replace '^\d{4}-\d{2}-\d{2}_','').ToLower()
})

if ($jsonFiles.Count -eq 0) {
    Write-Host ""
    Write-Host "  No projects found in $root" -ForegroundColor Yellow
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
            return
        }
    }
    return
}

$rows = @()
foreach ($jf in $jsonFiles) {
    try { $info = Get-Content $jf.FullName -Raw | ConvertFrom-Json } catch { continue }
    $rows += [PSCustomObject]@{
        Column                = if ($info.AnalyticsColumn)           { $info.AnalyticsColumn }                else { "-" }
        ColumnDescription     = if ($info.ColumnDescription)         { $info.ColumnDescription }              else { "-" }
        ProjectNo             = if ($null -ne $info.ProjectNo)       { [int]$info.ProjectNo }                 else { 0 }
        ProjectID             = if ($info.ProjectID)                 { $info.ProjectID }                      else { "-" }
        Project               = if ($info.Project)                   { $info.Project }                        else { "-" }
        PI                    = if ($info.PI)                        { $info.PI }                             else { "-" }
        TrapColumn            = if ($info.TrapColumn)                { $info.TrapColumn }                     else { "-" }
        TrapColumnDescription = if ($info.TrapColumnDescription)     { $info.TrapColumnDescription }          else { "-" }
        Created               = if ($info.Created)                   { $info.Created }                        else { "-" }
        SampleCount           = if ($info.SampleFolders)             { @($info.SampleFolders).Count }         else { 0 }
        SampleFolders         = if ($info.SampleFolders)             { (@($info.SampleFolders) -join "; ") }  else { "-" }
    }
}

$rows = @($rows | Sort-Object Column, ProjectNo)

# -- Load column library ------------------------------------------------------
$colLibFile = ".\data\columns.json"
$colLib = @()
if (Test-Path $colLibFile) {
    $loaded = Get-Content $colLibFile -Raw | ConvertFrom-Json
    if ($loaded) {
        $colLib = @($loaded)
        if ($colLib.Count -gt 0 -and $colLib[0] -is [PSCustomObject]) {
            $colLib = @($colLib | ForEach-Object { $_.Description })
        }
    }
}

# -- Filters ------------------------------------------------------------------
Write-Host ""
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  Filters  (blank = all)" -ForegroundColor Cyan
Write-Host ""

$fDateFrom = (Read-Host "  Date from (yyyy-MM-dd)").Trim()
$fDateTo   = (Read-Host "  Date to   (yyyy-MM-dd)").Trim()
$fPI       = (Read-Host "  PI").Trim()
Write-Host ""

# -- Column library picker ----------------------------------------------------
$fColumn = ""
if ($colLib.Count -gt 0) {
    Write-Host "  Column description  (Enter = select   Esc = no filter):" -ForegroundColor Cyan
    Write-Host ""

    $cLibSel = 0
    $cLibTop = [Console]::CursorTop

    for ($i = 0; $i -lt $colLib.Count; $i++) {
        [Console]::SetCursorPosition(0, $cLibTop + $i)
        $text = ("    " + $colLib[$i]).PadRight($w + 4)
        if ($i -eq $cLibSel) { Write-Host $text -ForegroundColor Black -BackgroundColor Cyan -NoNewline }
        else                  { Write-Host $text -ForegroundColor White -NoNewline }
    }
    [Console]::SetCursorPosition(0, $cLibTop + $colLib.Count)

    :colLibLoop while ($true) {
        $ck = [Console]::ReadKey($true)
        if ($ck.Key -eq [ConsoleKey]::UpArrow -or $ck.Key -eq [ConsoleKey]::DownArrow) {
            $prev    = $cLibSel
            $cLibSel = if ($ck.Key -eq [ConsoleKey]::UpArrow) {
                           ($cLibSel - 1 + $colLib.Count) % $colLib.Count
                       } else {
                           ($cLibSel + 1) % $colLib.Count
                       }
            [Console]::SetCursorPosition(0, $cLibTop + $prev)
            Write-Host ("    " + $colLib[$prev]).PadRight($w + 4) -ForegroundColor White -NoNewline
            [Console]::SetCursorPosition(0, $cLibTop + $cLibSel)
            Write-Host ("    " + $colLib[$cLibSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
            [Console]::SetCursorPosition(0, $cLibTop + $colLib.Count)
        } elseif ($ck.Key -eq [ConsoleKey]::Enter) {
            $fColumn = $colLib[$cLibSel]
            break colLibLoop
        } elseif ($ck.Key -eq [ConsoleKey]::Escape) {
            break colLibLoop
        }
    }
    Write-Host ""
} else {
    $fColumn = (Read-Host "  Column ID or description").Trim()
    Write-Host ""
}

# Parse and validate date inputs
$dtFrom = $null; $dtTo = $null
if ($fDateFrom -ne "") {
    try   { $dtFrom = [datetime]::ParseExact($fDateFrom, "yyyy-MM-dd", $null) }
    catch { Write-Host "  Invalid 'from' date - ignored." -ForegroundColor Yellow; $fDateFrom = "" }
}
if ($fDateTo -ne "") {
    try   { $dtTo = [datetime]::ParseExact($fDateTo, "yyyy-MM-dd", $null) }
    catch { Write-Host "  Invalid 'to' date - ignored." -ForegroundColor Yellow; $fDateTo = "" }
}

# Apply filters
if ($dtFrom) {
    $rows = @($rows | Where-Object {
        if ($_.Created -eq "-") { return $false }
        try { [datetime]::ParseExact($_.Created.Substring(0,10), "yyyy-MM-dd", $null) -ge $dtFrom } catch { $false }
    })
}
if ($dtTo) {
    $rows = @($rows | Where-Object {
        if ($_.Created -eq "-") { return $false }
        try { [datetime]::ParseExact($_.Created.Substring(0,10), "yyyy-MM-dd", $null) -le $dtTo } catch { $false }
    })
}
if ($fPI -ne "") {
    $rows = @($rows | Where-Object { $_.PI -like "*$fPI*" })
}
if ($fColumn -ne "") {
    $rows = @($rows | Where-Object {
        $_.Column -like "*$fColumn*" -or $_.ColumnDescription -like "*$fColumn*"
    })
}

# Active filter summary
$activeFilters = @()
if ($fDateFrom -ne "") { $activeFilters += "from $fDateFrom" }
if ($fDateTo   -ne "") { $activeFilters += "to $fDateTo" }
if ($fPI       -ne "") { $activeFilters += "PI=$fPI" }
if ($fColumn   -ne "") { $activeFilters += "column: $fColumn" }
$filterLabel = if ($activeFilters.Count -gt 0) { $activeFilters -join "  " } else { "none" }
Write-Host "  Filters: $filterLabel" -ForegroundColor DarkCyan
Write-Host ""

if ($rows.Count -eq 0) {
    Write-Host "  No projects match the applied filters." -ForegroundColor Yellow
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
            return
        }
    }
    return
}

# -- Zero-pad ProjectNo -------------------------------------------------------
$maxNo  = ($rows | Measure-Object ProjectNo -Maximum).Maximum
$digits = [Math]::Max(($maxNo.ToString().Length), 2)

# -- Console display ----------------------------------------------------------
$grouped = $rows | Group-Object Column
foreach ($grp in $grouped) {
    Write-Host "  $rule" -ForegroundColor DarkCyan
    Write-Host "  Column: $($grp.Name)  ($($grp.Count) project(s))" -ForegroundColor Cyan
    foreach ($r in $grp.Group) {
        $noStr = $r.ProjectNo.ToString().PadLeft($digits, '0')
        $proj  = if ($r.Project.Length -gt 24) { $r.Project.Substring(0,23) + "~" } else { $r.Project }
        Write-Host ("  #{0}  {1,-12} {2,-25} {3}" -f $noStr, $r.ProjectID, $proj, $r.PI)
    }
}
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host ""
Write-Host ("  {0} project(s) across {1} column(s)" -f $rows.Count, $grouped.Count) -ForegroundColor Cyan
Write-Host ""

# -- Export (zero-padded ProjectNo as string) ---------------------------------
$exportRows = $rows | ForEach-Object {
    [PSCustomObject]@{
        Column                = $_.Column
        ColumnDescription     = $_.ColumnDescription
        ProjectNo             = $_.ProjectNo.ToString().PadLeft($digits, '0')
        ProjectID             = $_.ProjectID
        Project               = $_.Project
        PI                    = $_.PI
        TrapColumn            = $_.TrapColumn
        TrapColumnDescription = $_.TrapColumnDescription
        Created               = $_.Created
        SampleCount           = $_.SampleCount
        SampleFolders         = $_.SampleFolders
    }
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outFile   = $null
if ($hasImportExcel) {
    $outFile = "$root\Projects_overview_$timestamp.xlsx"
    $exportRows | Export-Excel -Path $outFile -WorksheetName "Projects" `
        -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter `
        -TableName "ProjectsOverview" -TableStyle Medium6
    Write-Host "  Excel : $outFile" -ForegroundColor DarkCyan
} else {
    $outFile = "$root\Projects_overview_$timestamp.csv"
    $exportRows | Export-Csv -Path $outFile -NoTypeInformation
    Write-Host "  CSV   : $outFile" -ForegroundColor DarkCyan
}
Write-Host ""

# -- Navigation ---------------------------------------------------------------
$nItems = @("Open file", "Back to main menu", "Exit")
$nSel   = 0
$nTop   = [Console]::CursorTop

for ($i = 0; $i -lt $nItems.Count; $i++) {
    [Console]::SetCursorPosition(0, $nTop + $i)
    $fg = if ($nItems[$i] -eq "Exit") { "DarkYellow" } else { "DarkCyan" }
    Write-Host ("    " + $nItems[$i]).PadRight($w + 4) -ForegroundColor $fg -NoNewline
}
[Console]::SetCursorPosition(0, $nTop)
Write-Host ("  > " + $nItems[0]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
[Console]::SetCursorPosition(0, $nTop + $nItems.Count)

while ($true) {
    $k = [Console]::ReadKey($true)
    if ($k.Key -eq [ConsoleKey]::UpArrow -or $k.Key -eq [ConsoleKey]::DownArrow) {
        $p    = $nSel
        $nSel = if ($k.Key -eq [ConsoleKey]::UpArrow) {
                    ($nSel - 1 + $nItems.Count) % $nItems.Count
                } else {
                    ($nSel + 1) % $nItems.Count
                }
        [Console]::SetCursorPosition(0, $nTop + $p)
        $fg = if ($nItems[$p] -eq "Exit") { "DarkYellow" } else { "DarkCyan" }
        Write-Host ("    " + $nItems[$p]).PadRight($w + 4) -ForegroundColor $fg -NoNewline
        [Console]::SetCursorPosition(0, $nTop + $nSel)
        Write-Host ("  > " + $nItems[$nSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
        [Console]::SetCursorPosition(0, $nTop + $nItems.Count)
    } elseif ($k.Key -eq [ConsoleKey]::Enter) {
        if ($nItems[$nSel] -eq "Open file") {
            Invoke-Item $outFile
        } elseif ($nItems[$nSel] -eq "Back to main menu") {
            Clear-Host; .\Main.ps1; return
        } else {
            [Console]::SetCursorPosition(0, $nTop + $nItems.Count + 1)
            Write-Host "  Exiting..." -ForegroundColor DarkYellow
            return
        }
    } elseif ($k.Key -eq [ConsoleKey]::Escape) {
        [Console]::SetCursorPosition(0, $nTop + $nItems.Count + 1)
        Write-Host "  Exiting..." -ForegroundColor DarkYellow
        return
    }
}
