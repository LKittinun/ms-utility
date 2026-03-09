$w          = 55
$border     = "=" * $w
$rule       = "-" * $w
$prohibited = @("blank", "raw_summary", "prtc", "sst", "column_usage_history")

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "  [10]  Repair project order" -ForegroundColor Cyan
Write-Host "        Re-numbers projects by creation date" -ForegroundColor DarkCyan
Write-Host "        and rebuilds column_log.csv" -ForegroundColor DarkCyan
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host ""

# -- Password -----------------------------------------------------------------
$pwSS    = Read-Host "  Password" -AsSecureString
$pwPlain = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
               [Runtime.InteropServices.Marshal]::SecureStringToBSTR($pwSS))
$pwHash  = [System.BitConverter]::ToString(
               [System.Security.Cryptography.SHA256]::Create().ComputeHash(
                   [System.Text.Encoding]::UTF8.GetBytes($pwPlain)
               )).Replace("-","").ToLower()
if ($pwHash -ne "15e2b0d3c33891ebb0f1ef609ec419420c20e320ce94c65fbc8c3312448eb225") {
    Write-Host "  Access denied." -ForegroundColor Red
    Write-Host ""
    return
}

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

# ── Root ──────────────────────────────────────────────────────────────────────
$_cfg         = if (Test-Path (Join-Path $PSScriptRoot "config.json")) { Get-Content (Join-Path $PSScriptRoot "config.json") -Raw | ConvertFrom-Json } else { $null }
$root         = if ($_cfg -and $_cfg.Root) { $_cfg.Root } else { "Z:\Proteomics" }
$projectsRoot = Join-Path $root "Projects"
Write-Host "  Root : $root" -ForegroundColor DarkGray

# ── Analytics column ──────────────────────────────────────────────────────────
Write-Host ""
$analyticsCol = Read-Host "  Analytics column number (e.g. C20531700)"
if ($analyticsCol -eq "") {
    Write-Host "  Analytics column number cannot be empty." -ForegroundColor Red
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

# ── Resolve analytics column folder (date-prefixed) ──────────────────────────
$analyticsPath = $null
if (Test-Path $projectsRoot) {
    $existingColDir = Get-ChildItem $projectsRoot -Directory |
        Where-Object { $_.Name -like "*_$analyticsCol" } |
        Select-Object -First 1
    if ($existingColDir) { $analyticsPath = $existingColDir.FullName }
}
if (-not $analyticsPath) { $analyticsPath = Join-Path $projectsRoot $analyticsCol }
$logFile = Join-Path $analyticsPath "column_log.csv"

if (-not (Test-Path $analyticsPath)) {
    Write-Host "  Path not found: $analyticsPath" -ForegroundColor Red
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

# ── Scan projects ─────────────────────────────────────────────────────────────
$projects = Get-ChildItem -Path $analyticsPath -Directory |
    Where-Object { $prohibited -notcontains ($_.Name -replace '^\d{4}-\d{2}-\d{2}_','').ToLower() } |
    Where-Object { Test-Path (Join-Path $_.FullName "project_info.json") } |
    ForEach-Object {
        $json = Get-Content (Join-Path $_.FullName "project_info.json") -Raw | ConvertFrom-Json
        [PSCustomObject]@{
            Folder    = $_.FullName
            Name      = $_.Name
            Created   = [datetime]::ParseExact($json.Created, "yyyy-MM-dd HH:mm", $null)
            OldNo     = $json.ProjectNo
            Json      = $json
        }
    } |
    Sort-Object Created

if ($projects.Count -eq 0) {
    Write-Host "  No projects with project_info.json found in: $analyticsPath" -ForegroundColor Yellow
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

# ── Preview ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "  Proposed re-numbering (sorted by creation date):" -ForegroundColor Cyan
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host ("  " + "No.".PadRight(6) + "Was".PadRight(6) + "Created".PadRight(18) + "Project") -ForegroundColor DarkCyan

$newNo = 1
foreach ($p in $projects) {
    $changed = $p.OldNo -ne $newNo
    $color   = if ($changed) { "Yellow" } else { "White" }
    $marker  = if ($changed) { " *" } else { "" }
    Write-Host ("  " + "$newNo".PadRight(6) + "$($p.OldNo)".PadRight(6) + $p.Created.ToString("yyyy-MM-dd HH:mm").PadRight(18) + $p.Name + $marker) -ForegroundColor $color
    $newNo++
}
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  * = number will change" -ForegroundColor DarkCyan
Write-Host ""

# ── Confirm ───────────────────────────────────────────────────────────────────
$cItems = @("Yes, apply changes", "No, cancel")
$cSel   = 0
$cTop   = [Console]::CursorTop
[Console]::SetCursorPosition(0, $cTop);     Write-Host ("  > " + $cItems[0]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
[Console]::SetCursorPosition(0, $cTop + 1); Write-Host ("    " + $cItems[1]).PadRight($w + 4) -ForegroundColor DarkYellow -NoNewline
[Console]::SetCursorPosition(0, $cTop + 2)
while ($true) {
    $k = [Console]::ReadKey($true)
    if ($k.Key -eq [ConsoleKey]::UpArrow -or $k.Key -eq [ConsoleKey]::DownArrow) {
        $p = $cSel; $cSel = 1 - $cSel
        [Console]::SetCursorPosition(0, $cTop + $p);    Write-Host ("    " + $cItems[$p]).PadRight($w + 4) -ForegroundColor $(if ($p -eq 0) { "Cyan" } else { "DarkYellow" }) -NoNewline
        [Console]::SetCursorPosition(0, $cTop + $cSel); Write-Host ("  > " + $cItems[$cSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
    } elseif ($k.Key -eq [ConsoleKey]::Enter -or $k.Key -eq [ConsoleKey]::Escape) {
        [Console]::SetCursorPosition(0, $cTop + 2)
        if ($k.Key -ne [ConsoleKey]::Escape -and $cSel -eq 0) { break }
        Write-Host "  Cancelled." -ForegroundColor DarkYellow
        Write-Host ""
        Write-Host "  $rule" -ForegroundColor DarkCyan
        $nItems = @("Back to main menu", "Exit"); $nSel = 0
        $nTop = [Console]::CursorTop
        [Console]::SetCursorPosition(0, $nTop);     Write-Host ("  > " + $nItems[0]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
        [Console]::SetCursorPosition(0, $nTop + 1); Write-Host ("    " + $nItems[1]).PadRight($w + 4) -ForegroundColor DarkYellow -NoNewline
        [Console]::SetCursorPosition(0, $nTop + 2)
        while ($true) {
            $k2 = [Console]::ReadKey($true)
            if ($k2.Key -eq [ConsoleKey]::UpArrow -or $k2.Key -eq [ConsoleKey]::DownArrow) {
                $p = $nSel; $nSel = 1 - $nSel
                [Console]::SetCursorPosition(0, $nTop + $p);    Write-Host ("    " + $nItems[$p]).PadRight($w + 4) -ForegroundColor $(if ($p -eq 0) { "Cyan" } else { "DarkYellow" }) -NoNewline
                [Console]::SetCursorPosition(0, $nTop + $nSel); Write-Host ("  > " + $nItems[$nSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
            } elseif ($k2.Key -eq [ConsoleKey]::Enter -or $k2.Key -eq [ConsoleKey]::Escape) {
                if ($k2.Key -ne [ConsoleKey]::Escape -and $nSel -eq 0) { Clear-Host; .\Main.ps1 }
                else { [Console]::SetCursorPosition(0, $nTop + 3); Write-Host "  Exiting..." -ForegroundColor DarkYellow }
                return
            }
        }
        return
    }
}

# ── Apply ─────────────────────────────────────────────────────────────────────
Write-Host ""
$newLogRows = @()
$newNo = 1
foreach ($p in $projects) {
    $p.Json.ProjectNo = $newNo
    $p.Json | ConvertTo-Json | Out-File -FilePath (Join-Path $p.Folder "project_info.json") -Encoding UTF8
    $changed = if ($p.OldNo -ne $newNo) { " (was $($p.OldNo))" } else { "" }
    Write-Host "  [$newNo] $($p.Name)$changed" -ForegroundColor $(if ($p.OldNo -ne $newNo) { "Yellow" } else { "Green" })

    # Preserve existing ProjectID; generate one if missing (older projects)
    if (-not $p.Json.ProjectID) {
        $p.Json | Add-Member -NotePropertyName ProjectID -NotePropertyValue (
            -join ((65..90) + (48..57) | Get-Random -Count 8 | ForEach-Object { [char]$_ })
        )
    }

    $newLogRows += [PSCustomObject]@{
        ProjectID       = $p.Json.ProjectID
        ProjectNo       = $newNo
        Date            = $p.Json.Created
        Project         = $p.Json.Project
        PI              = if ($null -eq $p.Json.PI) { "" } else { $p.Json.PI }
        AnalyticsColumn = $p.Json.AnalyticsColumn
        TrapColumn      = if ($null -eq $p.Json.TrapColumn) { "" } else { $p.Json.TrapColumn }
        SampleFolders   = ($p.Json.SampleFolders -join ";")
    }
    $newNo++
}

# Rebuild column_log.csv from scratch
$newLogRows | Export-Csv $logFile -NoTypeInformation
Write-Host ""
Write-Host "  Rebuilt : $logFile" -ForegroundColor Green

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "  Done!  $($projects.Count) project(s) re-numbered." -ForegroundColor Cyan
Write-Host "  $border" -ForegroundColor DarkCyan

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
