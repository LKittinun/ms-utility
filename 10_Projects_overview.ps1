$w      = 55
$border = "=" * $w
$rule   = "-" * $w
$root   = "Z:\Proteomics\Projects"

Clear-Host
Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "  [10]  Projects overview" -ForegroundColor Cyan
Write-Host "        All projects across all columns (Excel)" -ForegroundColor DarkCyan
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

if ($jsonFiles.Count -eq 0) {
    Write-Host ""
    Write-Host "  No projects found in $root" -ForegroundColor Yellow
    Write-Host ""
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
    return
}

$rows = @()
foreach ($jf in $jsonFiles) {
    try { $info = Get-Content $jf.FullName -Raw | ConvertFrom-Json } catch { continue }

    $rows += [PSCustomObject]@{
        Column            = if ($info.AnalyticsColumn)       { $info.AnalyticsColumn }                    else { "-" }
        ColumnDescription = if ($info.ColumnDescription)     { $info.ColumnDescription }                  else { "-" }
        ProjectNo         = if ($null -ne $info.ProjectNo)   { [int]$info.ProjectNo }                     else { 0 }
        ProjectID         = if ($info.ProjectID)             { $info.ProjectID }                          else { "-" }
        Project           = if ($info.Project)               { $info.Project }                            else { "-" }
        PI                = if ($info.PI)                    { $info.PI }                                 else { "-" }
        TrapColumn        = if ($info.TrapColumn)            { $info.TrapColumn }                         else { "-" }
        TrapColumnDescription = if ($info.TrapColumnDescription) { $info.TrapColumnDescription }          else { "-" }
        Created           = if ($info.Created)               { $info.Created }                            else { "-" }
        SampleCount       = if ($info.SampleFolders)         { @($info.SampleFolders).Count }             else { 0 }
        SampleFolders     = if ($info.SampleFolders)         { (@($info.SampleFolders) -join "; ") }      else { "-" }
    }
}

$rows = $rows | Sort-Object Column, ProjectNo

# -- Console display ----------------------------------------------------------
Write-Host ""
$grouped = $rows | Group-Object Column
foreach ($grp in $grouped) {
    Write-Host "  $rule" -ForegroundColor DarkCyan
    Write-Host "  Column: $($grp.Name)  ($($grp.Count) project(s))" -ForegroundColor Cyan
    foreach ($r in $grp.Group) {
        $proj = if ($r.Project.Length -gt 26) { $r.Project.Substring(0,25) + "~" } else { $r.Project }
        $line = "  #{0,-4} {1,-12} {2,-27} {3}" -f $r.ProjectNo, $r.ProjectID, $proj, $r.PI
        Write-Host $line
    }
}
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host ""
Write-Host ("  {0} project(s) across {1} column(s)" -f $rows.Count, $grouped.Count) -ForegroundColor Cyan
Write-Host ""

# -- Export -------------------------------------------------------------------
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outFile   = $null

if ($hasImportExcel) {
    $outFile = "$root\Projects_overview_$timestamp.xlsx"
    $rows | Export-Excel -Path $outFile -WorksheetName "Projects" `
        -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter `
        -TableName "ProjectsOverview" -TableStyle Medium6
    Write-Host "  Excel : $outFile" -ForegroundColor DarkCyan
} else {
    $outFile = "$root\Projects_overview_$timestamp.csv"
    $rows | Export-Csv -Path $outFile -NoTypeInformation
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
