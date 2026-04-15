$w          = 55
$border     = "=" * $w
$rule       = "-" * $w
$prohibited = @("blank", "raw_summary", "prtc", "sst", "column_usage_history")

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "   [1]  Project folder initializer" -ForegroundColor Cyan
Write-Host "        Creates project structure and logs column info" -ForegroundColor DarkCyan
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

# ── Root ──────────────────────────────────────────────────────────────────────
$_cfg         = if (Test-Path (Join-Path $PSScriptRoot "config.json")) { Get-Content (Join-Path $PSScriptRoot "config.json") -Raw | ConvertFrom-Json } else { $null }
$root         = if ($_cfg -and $_cfg.Root) { $_cfg.Root } else { "Z:\Proteomics" }
$projectsRoot = Join-Path $root "Projects"
Write-Host "  Root : $root" -ForegroundColor DarkGray

# ── Column library ────────────────────────────────────────────────────────────
$colLibFile = ".\data\columns.json"
$colLib     = @()
if (Test-Path $colLibFile) {
    $loaded = Get-Content $colLibFile -Raw | ConvertFrom-Json
    if ($loaded) {
        $colLib = @($loaded)
        # Migrate old object format { ColumnID, Description } -> string array
        if ($colLib.Count -gt 0 -and $colLib[0] -is [PSCustomObject]) {
            $colLib = @($colLib | ForEach-Object { $_.Description })
            if (-not (Test-Path ".\data")) { [System.IO.Directory]::CreateDirectory(".\data") | Out-Null }
            ConvertTo-Json -InputObject @($colLib) | Out-File $colLibFile -Encoding UTF8
        }
    }
}

# ── Analytics column ID ───────────────────────────────────────────────────────
Write-Host ""
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  Analytics column" -ForegroundColor Cyan
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host ""

# Ask: same column (pick from list) or new?
$scItems = @("Yes - select from previous columns", "No - enter new column ID")
$scSel   = 0
$scTop   = [Console]::CursorTop
[Console]::SetCursorPosition(0, $scTop)
Write-Host ("  > " + $scItems[0]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
[Console]::SetCursorPosition(0, $scTop + 1)
Write-Host ("    " + $scItems[1]).PadRight($w + 4) -ForegroundColor DarkCyan -NoNewline
[Console]::SetCursorPosition(0, $scTop + 2)
while ($true) {
    $k = [Console]::ReadKey($true)
    if ($k.Key -eq [ConsoleKey]::UpArrow -or $k.Key -eq [ConsoleKey]::DownArrow) {
        $p = $scSel; $scSel = 1 - $scSel
        [Console]::SetCursorPosition(0, $scTop + $p)
        Write-Host ("    " + $scItems[$p]).PadRight($w + 4) -ForegroundColor DarkCyan -NoNewline
        [Console]::SetCursorPosition(0, $scTop + $scSel)
        Write-Host ("  > " + $scItems[$scSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
    } elseif ($k.Key -eq [ConsoleKey]::Enter) {
        break
    }
}
[Console]::SetCursorPosition(0, $scTop + $scItems.Count)
Write-Host ""

$analyticsCol      = ""
$colDesc           = ""
$colDescFromPicker = $false
if ($scSel -eq 0) {
    # Read (ColID, Description) from column_info.json in each column folder, most recent first
    $prevColPairs = @()
    if (Test-Path $projectsRoot) {
        $prevColPairs = @(
            Get-ChildItem $projectsRoot -Directory |
            Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}_' } |
            Sort-Object Name -Descending |
            ForEach-Object {
                $infoPath = Join-Path $_.FullName "column_info.json"
                if (Test-Path $infoPath) {
                    $ci   = Get-Content $infoPath -Raw | ConvertFrom-Json
                    $cid  = if ($ci.ColumnID)    { $ci.ColumnID }    else { $_.Name -replace '^\d{4}-\d{2}-\d{2}_', '' }
                    $cdsc = if ($ci.Description) { $ci.Description } else { "" }
                    $lbl  = if ($cdsc) { "$cid  [$cdsc]" } else { $cid }
                    [PSCustomObject]@{ ID = $cid; Desc = $cdsc; Label = $lbl }
                } else {
                    # No column_info.json - fall back to folder name
                    $cid = $_.Name -replace '^\d{4}-\d{2}-\d{2}_', ''
                    [PSCustomObject]@{ ID = $cid; Desc = ""; Label = $cid }
                }
            }
        )
    }
    if ($prevColPairs.Count -eq 0) {
        Write-Host "  No previous columns found. Enter new column ID." -ForegroundColor Yellow
        Write-Host ""
        $analyticsCol = Read-Host "  Column ID (e.g. C20533039, no date prefix)"
    } else {
        Write-Host "  Select column:" -ForegroundColor Cyan
        Write-Host "  (Up/Down: select   Enter: confirm)" -ForegroundColor DarkGray
        Write-Host ""

        $colIDItems = $prevColPairs
        $colIDSel   = 0
        $colIDTop   = [Console]::CursorTop

        function DrawColIDItem($idx, $hl) {
            [Console]::SetCursorPosition(0, $colIDTop + $idx)
            $text = ("    " + $colIDItems[$idx].Label).PadRight($w + 4)
            if ($hl) { Write-Host $text -ForegroundColor Black -BackgroundColor Cyan -NoNewline }
            else     { Write-Host $text -ForegroundColor White -NoNewline }
        }

        for ($i = 0; $i -lt $colIDItems.Count; $i++) {
            DrawColIDItem $i ($i -eq $colIDSel)
            [Console]::SetCursorPosition(0, $colIDTop + $i + 1)
        }
        [Console]::SetCursorPosition(0, $colIDTop + $colIDItems.Count + 1)

        while ($true) {
            $k = [Console]::ReadKey($true)
            if ($k.Key -eq [ConsoleKey]::UpArrow) {
                $prev = $colIDSel; $colIDSel = ($colIDSel - 1 + $colIDItems.Count) % $colIDItems.Count
                DrawColIDItem $prev $false; DrawColIDItem $colIDSel $true
            } elseif ($k.Key -eq [ConsoleKey]::DownArrow) {
                $prev = $colIDSel; $colIDSel = ($colIDSel + 1) % $colIDItems.Count
                DrawColIDItem $prev $false; DrawColIDItem $colIDSel $true
            } elseif ($k.Key -eq [ConsoleKey]::Enter) {
                [Console]::SetCursorPosition(0, $colIDTop + $colIDItems.Count + 1)
                $analyticsCol      = $colIDItems[$colIDSel].ID
                $colDesc           = $colIDItems[$colIDSel].Desc
                $colDescFromPicker = $true
                break
            }
        }
    }
} else {
    $analyticsCol = Read-Host "  Column ID (e.g. C20533039, no date prefix)"
}

if ($analyticsCol -eq "") {
    Write-Host "  Analytics column cannot be empty." -ForegroundColor Red
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

# ── Column description (library selection) ────────────────────────────────────
if (-not $colDescFromPicker) {
    Write-Host ""
    Write-Host "  Column description:" -ForegroundColor Cyan
    Write-Host "  (Up/Down: select   Del: remove from library   Enter: confirm)" -ForegroundColor DarkGray
    Write-Host ""
}

function DrawDescItem($idx, $hl) {
    [Console]::SetCursorPosition(0, $descTop + $idx)
    $isAdd = ($idx -eq $descMenuItems.Count - 1)
    if ($isAdd) {
        $text = "    [+ Add new description]".PadRight($w + 4)
        if ($hl) { Write-Host $text -ForegroundColor Black -BackgroundColor Cyan -NoNewline }
        else      { Write-Host $text -ForegroundColor DarkYellow -NoNewline }
    } else {
        $text = ("    " + $descMenuItems[$idx]).PadRight($w + 4)
        if ($hl) { Write-Host $text -ForegroundColor Black -BackgroundColor Cyan -NoNewline }
        else      { Write-Host $text -ForegroundColor White -NoNewline }
    }
}

if (-not $colDescFromPicker) { :descLoop while ($true) {
    if ($colLib.Count -eq 0) {
        $newDesc = Read-Host "  New description (leave blank to skip)"
        if ($newDesc -ne "") {
            $colLib += $newDesc
            if (-not (Test-Path ".\data")) { [System.IO.Directory]::CreateDirectory(".\data") | Out-Null }
            ConvertTo-Json -InputObject @($colLib) | Out-File $colLibFile -Encoding UTF8
            Write-Host "  Saved to description library." -ForegroundColor Green
        }
        $colDesc = $newDesc
        break
    }

    $descMenuItems = @($colLib) + @("[+ Add new description]")
    $descSel = 0
    $descTop = [Console]::CursorTop

    for ($i = 0; $i -lt $descMenuItems.Count; $i++) {
        DrawDescItem $i ($i -eq $descSel)
        [Console]::SetCursorPosition(0, $descTop + $i + 1)
    }
    [Console]::SetCursorPosition(0, $descTop + $descMenuItems.Count + 1)

    while ($true) {
        $k = [Console]::ReadKey($true)
        if ($k.Key -eq [ConsoleKey]::UpArrow) {
            $prev = $descSel; $descSel = ($descSel - 1 + $descMenuItems.Count) % $descMenuItems.Count
            DrawDescItem $prev $false; DrawDescItem $descSel $true
        } elseif ($k.Key -eq [ConsoleKey]::DownArrow) {
            $prev = $descSel; $descSel = ($descSel + 1) % $descMenuItems.Count
            DrawDescItem $prev $false; DrawDescItem $descSel $true
        } elseif ($k.Key -eq [ConsoleKey]::Delete) {
            $isAddItem = ($descSel -eq $descMenuItems.Count - 1)
            if (-not $isAddItem) {
                $toRemove = $descMenuItems[$descSel]
                $colLib = @($colLib | Where-Object { $_ -ne $toRemove })
                if (-not (Test-Path ".\data")) { [System.IO.Directory]::CreateDirectory(".\data") | Out-Null }
                ConvertTo-Json -InputObject @($colLib) | Out-File $colLibFile -Encoding UTF8
                [Console]::SetCursorPosition(0, $descTop + $descMenuItems.Count + 1)
                Write-Host ""
                continue descLoop
            }
        } elseif ($k.Key -eq [ConsoleKey]::Enter) {
            $isAddItem = ($descSel -eq $descMenuItems.Count - 1)
            [Console]::SetCursorPosition(0, $descTop + $descMenuItems.Count + 1)
            if ($isAddItem) {
                Write-Host ""
                $newDesc = Read-Host "  New description (leave blank to skip)"
                if ($newDesc -ne "") {
                    $colLib += $newDesc
                    if (-not (Test-Path ".\data")) { [System.IO.Directory]::CreateDirectory(".\data") | Out-Null }
                    ConvertTo-Json -InputObject @($colLib) | Out-File $colLibFile -Encoding UTF8
                    Write-Host "  Saved to description library." -ForegroundColor Green
                    continue descLoop
                } else {
                    $colDesc = ""
                    break
                }
            } else {
                $colDesc = $descMenuItems[$descSel]
                break
            }
        }
    }
    break
} } # end if (-not $colDescFromPicker)

# ── Resolve analytics column folder (date-prefixed, e.g. 2026-03-02_C20533039) ─
$colDatePrefix = Get-Date -Format "yyyy-MM-dd"
$analyticsPath = $null
if (Test-Path $projectsRoot) {
    $existingColDir = Get-ChildItem $projectsRoot -Directory |
        Where-Object { $_.Name -like "*_$analyticsCol" } |
        Select-Object -First 1
    if ($existingColDir) { $analyticsPath = $existingColDir.FullName }
}
if (-not $analyticsPath) { $analyticsPath = Join-Path $projectsRoot "${colDatePrefix}_${analyticsCol}" }
$logFile = Join-Path $analyticsPath "column_log.csv"

# ── Column info JSON ───────────────────────────────────────────────────────────
$colInfoFile    = Join-Path $analyticsPath "column_info.json"
$colInfoData    = $null
$colInfoChanged = $false
if (Test-Path $colInfoFile) {
    $colInfoData = Get-Content $colInfoFile -Raw | ConvertFrom-Json
}

if ($null -eq $colInfoData) {
    # New column - collect all fields
    Write-Host ""
    Write-Host "  $rule" -ForegroundColor DarkCyan
    Write-Host "  New column detected - enter column details (all optional):" -ForegroundColor Cyan
    Write-Host "  $rule" -ForegroundColor DarkCyan
    Write-Host ""
    $colFirstUse = Read-Host "  First use date (yyyy-MM-dd)"
    $colInfoChanged = $true
} else {
    # Existing column - carry over fields
    $colFirstUse = if ($colInfoData.FirstUseDate)  { $colInfoData.FirstUseDate }  else { "" }
    Write-Host ""
    Write-Host "  Column: $analyticsCol" -ForegroundColor Cyan
}

# ── Project selection ─────────────────────────────────────────────────────────
$existingInfo = $null
$projectPath  = $null

Write-Host ""
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  Project" -ForegroundColor Cyan
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host ""

$existingDirs = @()
if (Test-Path $analyticsPath) {
    $existingDirs = @(
        Get-ChildItem $analyticsPath -Directory |
        Where-Object { Test-Path (Join-Path $_.FullName "project_info.json") } |
        Sort-Object Name -Descending
    )
}

$projIDItems = @("[ New project ]") + @($existingDirs | ForEach-Object {
    $j   = Get-Content (Join-Path $_.FullName "project_info.json") -Raw | ConvertFrom-Json
    $lbl = $j.Project
    if ($j.PI) { $lbl += "  ($($j.PI))" }
    $lbl
})
$projSel = 0
Write-Host "  (Up/Down: select   Enter: confirm)" -ForegroundColor DarkGray
Write-Host ""
$projTop = [Console]::CursorTop

function DrawProjItem($idx, $hl) {
    [Console]::SetCursorPosition(0, $projTop + $idx)
    $text = ("    " + $projIDItems[$idx]).PadRight($w + 4)
    if ($hl)           { Write-Host $text -ForegroundColor Black -BackgroundColor Cyan -NoNewline }
    elseif ($idx -eq 0){ Write-Host $text -ForegroundColor DarkYellow -NoNewline }
    else               { Write-Host $text -ForegroundColor White -NoNewline }
}

for ($i = 0; $i -lt $projIDItems.Count; $i++) {
    DrawProjItem $i ($i -eq $projSel)
    [Console]::SetCursorPosition(0, $projTop + $i + 1)
}
[Console]::SetCursorPosition(0, $projTop + $projIDItems.Count + 1)

while ($true) {
    $k = [Console]::ReadKey($true)
    if ($k.Key -eq [ConsoleKey]::UpArrow) {
        $prev = $projSel; $projSel = ($projSel - 1 + $projIDItems.Count) % $projIDItems.Count
        DrawProjItem $prev $false; DrawProjItem $projSel $true
    } elseif ($k.Key -eq [ConsoleKey]::DownArrow) {
        $prev = $projSel; $projSel = ($projSel + 1) % $projIDItems.Count
        DrawProjItem $prev $false; DrawProjItem $projSel $true
    } elseif ($k.Key -eq [ConsoleKey]::Enter) {
        [Console]::SetCursorPosition(0, $projTop + $projIDItems.Count)
        Write-Host ""
        break
    }
}

$projectName = ""
if ($projSel -eq 0) {
    # New project - prompt for name
    $projectName = Read-Host "  Project name"
    if ($projectName -eq "") {
        Write-Host "  Project name cannot be empty." -ForegroundColor Red
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
    $projectName = $projectName -replace '[\s]+', '_' -replace '[<>:"/\\|?*]', ''
    if ($projectName -eq "") {
        Write-Host "  Project name is empty after sanitizing invalid characters." -ForegroundColor Red
        return
    }
    if ($prohibited -contains $projectName.ToLower()) {
        Write-Host "  '$projectName' is a reserved name and cannot be used as a project name." -ForegroundColor Red
        return
    }
    $datePrefix  = Get-Date -Format "yyyy-MM-dd"
    $projectPath = Join-Path $analyticsPath "${datePrefix}_${projectName}"
    if (Test-Path $projectPath) {
        Write-Host "  WARNING: project folder already exists." -ForegroundColor Yellow
    }
} else {
    # Existing project - load info, no further prompts for fixed fields
    $chosenDir    = $existingDirs[$projSel - 1]
    $projectPath  = $chosenDir.FullName
    $existingInfo = Get-Content (Join-Path $projectPath "project_info.json") -Raw | ConvertFrom-Json
    $projectName  = $existingInfo.Project
    Write-Host "  Existing project: $projectName" -ForegroundColor Yellow
    Write-Host "  (leave fields blank to keep current values)" -ForegroundColor DarkGray
}

# ── PI ────────────────────────────────────────────────────────────────────────
Write-Host ""
if ($existingInfo) {
    $pi = if ($existingInfo.PI) { $existingInfo.PI } else { "" }
    Write-Host "  PI: $(if ($pi) { $pi } else { '(not specified)' })" -ForegroundColor Cyan
} else {
    $piRaw = Read-Host "  PI name (leave blank if unknown)"
    $pi    = $piRaw
}

# ── Trap column ───────────────────────────────────────────────────────────────
$trapLibFile = ".\data\trap_columns.json"
$trapLib     = @()
if (Test-Path $trapLibFile) {
    $tLoaded = Get-Content $trapLibFile -Raw | ConvertFrom-Json
    if ($tLoaded) { $trapLib = @($tLoaded) }
}

# Step 1: Collect unique (TrapColumn, Description) pairs from this analytics column's projects
$trapColPairs = @()
if (Test-Path $analyticsPath) {
    $tSeen     = [System.Collections.Generic.HashSet[string]]::new()
    $tPairList = [System.Collections.Generic.List[object]]::new()
    Get-ChildItem $analyticsPath -Recurse -Filter "project_info.json" -File |
    Sort-Object DirectoryName -Descending |
    ForEach-Object {
        $j    = Get-Content $_.FullName -Raw | ConvertFrom-Json
        $tid  = $j.TrapColumn
        $tdsc = if ($j.TrapColumnDescription) { $j.TrapColumnDescription } else { "" }
        if ($tid) {
            $key = "$tid|$tdsc"
            if ($tSeen.Add($key)) {
                $lbl = if ($tdsc) { "$tid  [$tdsc]" } else { $tid }
                $tPairList.Add([PSCustomObject]@{ ID = $tid; Desc = $tdsc; Label = $lbl })
            }
        }
    }
    $trapColPairs = @($tPairList)
}

Write-Host ""
Write-Host "  Trap column ID:" -ForegroundColor Cyan
if ($existingInfo) { Write-Host "  (current: $(if ($existingInfo.TrapColumn) { $existingInfo.TrapColumn } else { '(not specified)' }))" -ForegroundColor DarkGray }
Write-Host "  (Up/Down: select   Enter: confirm)" -ForegroundColor DarkGray
Write-Host ""

$trapIDItems = @("[ No trap column ]") + @($trapColPairs | ForEach-Object { $_.Label }) + @("[+ Type new ID]")
$trapIDSel   = 0
if ($existingInfo -and $existingInfo.TrapColumn) {
    # Pre-select matching pair (same ID and description)
    $preLabel = if ($existingInfo.TrapColumnDescription) { "$($existingInfo.TrapColumn)  [$($existingInfo.TrapColumnDescription)]" } else { $existingInfo.TrapColumn }
    $preIdx   = [array]::IndexOf($trapIDItems, $preLabel)
    if ($preIdx -lt 0) { $preIdx = [array]::IndexOf($trapIDItems, $existingInfo.TrapColumn) }
    if ($preIdx -ge 0) { $trapIDSel = $preIdx }
}
$trapIDTop = [Console]::CursorTop

function DrawTrapIDItem($idx, $hl) {
    [Console]::SetCursorPosition(0, $trapIDTop + $idx)
    $isSpecial = ($idx -eq 0 -or $idx -eq $trapIDItems.Count - 1)
    $text      = ("    " + $trapIDItems[$idx]).PadRight($w + 4)
    if ($hl)            { Write-Host $text -ForegroundColor Black -BackgroundColor Cyan -NoNewline }
    elseif ($isSpecial) { Write-Host $text -ForegroundColor DarkYellow -NoNewline }
    else                { Write-Host $text -ForegroundColor White -NoNewline }
}

for ($i = 0; $i -lt $trapIDItems.Count; $i++) {
    DrawTrapIDItem $i ($i -eq $trapIDSel)
    [Console]::SetCursorPosition(0, $trapIDTop + $i + 1)
}
[Console]::SetCursorPosition(0, $trapIDTop + $trapIDItems.Count + 1)

$trapCol          = ""
$trapColDesc      = ""
$trapDescFromPicker = $false
while ($true) {
    $k = [Console]::ReadKey($true)
    if ($k.Key -eq [ConsoleKey]::UpArrow) {
        $prev = $trapIDSel; $trapIDSel = ($trapIDSel - 1 + $trapIDItems.Count) % $trapIDItems.Count
        DrawTrapIDItem $prev $false; DrawTrapIDItem $trapIDSel $true
    } elseif ($k.Key -eq [ConsoleKey]::DownArrow) {
        $prev = $trapIDSel; $trapIDSel = ($trapIDSel + 1) % $trapIDItems.Count
        DrawTrapIDItem $prev $false; DrawTrapIDItem $trapIDSel $true
    } elseif ($k.Key -eq [ConsoleKey]::Enter) {
        [Console]::SetCursorPosition(0, $trapIDTop + $trapIDItems.Count + 1)
        if ($trapIDSel -eq 0) {
            $trapCol = ""
        } elseif ($trapIDSel -eq $trapIDItems.Count - 1) {
            Write-Host ""
            $trapCol = Read-Host "  Trap column ID"
        } else {
            $pair               = $trapColPairs[$trapIDSel - 1]
            $trapCol            = $pair.ID
            $trapColDesc        = $pair.Desc
            $trapDescFromPicker = $true
        }
        break
    }
}

# Step 2: Trap column description - scoped to descriptions recorded for the selected ID
if ($trapCol -ne "" -and -not $trapDescFromPicker) {
    # Collect descriptions already used with this specific trap column ID
    $trapColDescs = @()
    if (Test-Path $analyticsPath) {
        $trapColDescs = @(
            Get-ChildItem $analyticsPath -Recurse -Filter "project_info.json" -File |
            ForEach-Object {
                $j = Get-Content $_.FullName -Raw | ConvertFrom-Json
                if ($j.TrapColumn -eq $trapCol) { $j.TrapColumnDescription }
            } |
            Where-Object { $_ -and $_ -ne "" } |
            Sort-Object -Unique
        )
    }

    Write-Host ""
    Write-Host "  Trap column description:" -ForegroundColor Cyan
    Write-Host "  (Up/Down: select   Del: hide from list   Enter: confirm)" -ForegroundColor DarkGray
    Write-Host ""

    :trapDescLoop while ($true) {
        if ($trapColDescs.Count -eq 0) {
            $newDesc = Read-Host "  New description (leave blank to skip)"
            if ($newDesc -ne "") {
                $trapLib += $newDesc
                if (-not (Test-Path ".\data")) { [System.IO.Directory]::CreateDirectory(".\data") | Out-Null }
                ConvertTo-Json -InputObject @($trapLib) | Out-File $trapLibFile -Encoding UTF8
                Write-Host "  Saved to trap column description library." -ForegroundColor Green
            }
            $trapColDesc = $newDesc
            break
        }

        $descMenuItems = @($trapColDescs) + @("[+ Add new description]")
        $descSel = 0
        if ($existingInfo -and $existingInfo.TrapColumnDescription) {
            $preIdx = [array]::IndexOf($trapColDescs, $existingInfo.TrapColumnDescription)
            if ($preIdx -ge 0) { $descSel = $preIdx }
        }
        $descTop = [Console]::CursorTop

        for ($i = 0; $i -lt $descMenuItems.Count; $i++) {
            DrawDescItem $i ($i -eq $descSel)
            [Console]::SetCursorPosition(0, $descTop + $i + 1)
        }
        [Console]::SetCursorPosition(0, $descTop + $descMenuItems.Count + 1)

        while ($true) {
            $k = [Console]::ReadKey($true)
            if ($k.Key -eq [ConsoleKey]::UpArrow) {
                $prev = $descSel; $descSel = ($descSel - 1 + $descMenuItems.Count) % $descMenuItems.Count
                DrawDescItem $prev $false; DrawDescItem $descSel $true
            } elseif ($k.Key -eq [ConsoleKey]::DownArrow) {
                $prev = $descSel; $descSel = ($descSel + 1) % $descMenuItems.Count
                DrawDescItem $prev $false; DrawDescItem $descSel $true
            } elseif ($k.Key -eq [ConsoleKey]::Delete) {
                $isAddItem = ($descSel -eq $descMenuItems.Count - 1)
                if (-not $isAddItem) {
                    $toRemove = $descMenuItems[$descSel]
                    $trapColDescs = @($trapColDescs | Where-Object { $_ -ne $toRemove })
                    [Console]::SetCursorPosition(0, $descTop + $descMenuItems.Count + 1)
                    Write-Host ""
                    continue trapDescLoop
                }
            } elseif ($k.Key -eq [ConsoleKey]::Enter) {
                $isAddItem = ($descSel -eq $descMenuItems.Count - 1)
                [Console]::SetCursorPosition(0, $descTop + $descMenuItems.Count + 1)
                if ($isAddItem) {
                    Write-Host ""
                    $newDesc = Read-Host "  New description (leave blank to skip)"
                    if ($newDesc -ne "") {
                        $trapColDescs += $newDesc
                        $trapLib += $newDesc
                        if (-not (Test-Path ".\data")) { [System.IO.Directory]::CreateDirectory(".\data") | Out-Null }
                        ConvertTo-Json -InputObject @($trapLib) | Out-File $trapLibFile -Encoding UTF8
                        Write-Host "  Saved to trap column description library." -ForegroundColor Green
                        continue trapDescLoop
                    } else {
                        $trapColDesc = ""
                        break
                    }
                } else {
                    $trapColDesc = $descMenuItems[$descSel]
                    break
                }
            }
        }
        break
    }
}

# ── Sample subfolders ─────────────────────────────────────────────────────────
Write-Host ""
$existingFolders = if ($existingInfo -and $existingInfo.SampleFolders) { @($existingInfo.SampleFolders) } else { @() }
if ($existingFolders.Count -gt 0) {
    Write-Host "  (current: $($existingFolders -join ', '))" -ForegroundColor DarkGray
}
$subfolders = @()
while ($true) {
    $subfoldersRaw = Read-Host "  Sample subfolders, comma-separated$(if ($existingInfo) { ' (blank to keep, type new to add)' } else { ' (e.g. Plasma,pEV)' })"
    if ($existingInfo -and $subfoldersRaw -eq "") {
        $subfolders = $existingFolders
        break
    }
    $parsed = @($subfoldersRaw -split "," | ForEach-Object { ($_.Trim() -replace '[\s]+', '_' -replace '[<>:"/\\|?*]', '') } | Where-Object { $_ -ne "" })
    # Reserved names
    $reservedHits = @($parsed | Where-Object { $prohibited -contains $_.ToLower() })
    if ($reservedHits.Count -gt 0) {
        Write-Host "  Reserved names not allowed as subfolders: $($reservedHits -join ', ')" -ForegroundColor Red
        continue
    }
    # Duplicates within new input
    $seen  = @{}
    $dupes = @()
    foreach ($sf in $parsed) {
        $key = $sf.ToLower()
        if ($seen.ContainsKey($key)) { if ($dupes -notcontains $sf) { $dupes += $sf } }
        else { $seen[$key] = $true }
    }
    if ($dupes.Count -gt 0) {
        Write-Host "  Duplicate subfolders not allowed: $($dupes -join ', ')" -ForegroundColor Red
        continue
    }
    # Conflicts with already-existing subfolders in JSON
    $conflicts = @($parsed | Where-Object { $existingFolders -icontains $_ })
    if ($conflicts.Count -gt 0) {
        Write-Host "  Already exist in this project: $($conflicts -join ', ')" -ForegroundColor Red
        continue
    }
    $subfolders = @($existingFolders) + @($parsed)
    break
}

# ── Project number + ID ───────────────────────────────────────────────────────
if ($existingInfo) {
    $projectID = $existingInfo.ProjectID
    $projectNo = $existingInfo.ProjectNo
} else {
    $projectNo = 1
    if (Test-Path $analyticsPath) {
        $existingProjects = (Get-ChildItem -Path $analyticsPath -Directory |
            Where-Object { Test-Path (Join-Path $_.FullName "project_info.json") }).Count
        $projectNo = $existingProjects + 1
    }
    $projectID = -join ((65..90) + (48..57) | Get-Random -Count 8 | ForEach-Object { [char]$_ })
}

# ── Preview (tree) ────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  Folders to create:" -ForegroundColor Cyan
Write-Host "  Projects\$(Split-Path $analyticsPath -Leaf)\" -ForegroundColor DarkGray
Write-Host "  \-- $(Split-Path $projectPath -Leaf)\" -ForegroundColor White
if ($subfolders.Count -eq 0) {
    Write-Host "      \-- Result\" -ForegroundColor Gray
} else {
    for ($i = 0; $i -lt $subfolders.Count; $i++) {
        $isLast = ($i -eq $subfolders.Count - 1)
        $branch = if ($isLast) { "\--" } else { "+--" }
        $pipe   = if ($isLast) { "    " } else { "|   " }
        Write-Host "      $branch $($subfolders[$i])\" -ForegroundColor White
        Write-Host "      $pipe\-- Result\" -ForegroundColor Gray
    }
}
Write-Host ""
Write-Host "  Column log entry (project no. $projectNo):" -ForegroundColor Cyan
Write-Host "    ID        : $projectID" -ForegroundColor White
Write-Host "    PI        : $(if ($pi -eq '') { '(not specified)' } else { $pi })" -ForegroundColor White
Write-Host "    Analytics : $analyticsCol" -ForegroundColor White
Write-Host "    Desc      : $(if ($colDesc -eq '') { '(none)' } else { $colDesc })" -ForegroundColor DarkGray
Write-Host "    Trap      : $(if ($trapCol -eq '') { '(not specified)' } else { $trapCol })" -ForegroundColor White
Write-Host "    TrapDesc  : $(if ($trapColDesc -eq '') { '(none)' } else { $trapColDesc })" -ForegroundColor DarkGray
Write-Host "    Project   : $projectName" -ForegroundColor White
if ($colInfoChanged) {
    Write-Host ""
    Write-Host "  New column_info.json:" -ForegroundColor Cyan
    Write-Host "    First use date: $(if ($colFirstUse -eq '') { '(not specified)' } else { $colFirstUse })" -ForegroundColor White
}
$rawCount = @(Get-ChildItem -Path $projectPath -Recurse -Filter "*.raw" -File -ErrorAction SilentlyContinue).Count
if ($rawCount -gt 0) {
    Write-Host ""
    Write-Host "  Note: $rawCount .raw file(s) exist in this folder - they will not be affected." -ForegroundColor DarkYellow
}
Write-Host ""

Write-Host "  $rule" -ForegroundColor DarkCyan
$cItems = @("Yes, create folders", "No, cancel")
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

# ── Create folders ────────────────────────────────────────────────────────────
Write-Host ""
try {
    [System.IO.Directory]::CreateDirectory($projectPath) | Out-Null
    if ($subfolders.Count -eq 0) {
        [System.IO.Directory]::CreateDirectory("$projectPath\Result") | Out-Null
    } else {
        foreach ($sf in $subfolders) {
            [System.IO.Directory]::CreateDirectory("$projectPath\$sf")         | Out-Null
            [System.IO.Directory]::CreateDirectory("$projectPath\$sf\Result") | Out-Null
        }
    }
} catch {
    Write-Host "  ERROR creating folders: $_" -ForegroundColor Red
    Write-Host "  No files were written." -ForegroundColor Red
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

# ── Write project_info.json ───────────────────────────────────────────────────
$now = Get-Date -Format "yyyy-MM-dd HH:mm"
[PSCustomObject]@{
    ProjectID         = $projectID
    Project           = $projectName
    PI                = if ($pi -eq "") { $null } else { $pi }
    AnalyticsColumn   = $analyticsCol
    ColumnDescription = if ($colDesc -eq "") { $null } else { $colDesc }
    TrapColumn            = if ($trapCol -eq "") { $null } else { $trapCol }
    TrapColumnDescription = if ($trapColDesc -eq "") { $null } else { $trapColDesc }
    ProjectNo         = $projectNo
    Created           = $now
    SampleFolders     = @($subfolders)
} | ConvertTo-Json | Out-File -FilePath "$projectPath\project_info.json" -Encoding UTF8

# ── Write column_info.json ────────────────────────────────────────────────────
if ($colInfoChanged) {
    [System.IO.Directory]::CreateDirectory($analyticsPath) | Out-Null
    [PSCustomObject]@{
        ColumnID     = $analyticsCol
        Description  = if ($colDesc -eq "") { $null } else { $colDesc }
        FirstUseDate = if ($colFirstUse -eq "") { $null } else { $colFirstUse }
        Created      = (Get-Date -Format "yyyy-MM-dd HH:mm")
    } | ConvertTo-Json | Out-File $colInfoFile -Encoding UTF8
}

# ── Append to column_log.csv ──────────────────────────────────────────────────
$logRow = [PSCustomObject]@{
    ProjectID         = $projectID
    ProjectNo         = $projectNo
    Date              = $now
    Project           = $projectName
    PI                = $pi
    AnalyticsColumn   = $analyticsCol
    ColumnDescription = $colDesc
    TrapColumn            = $trapCol
    TrapColumnDescription = $trapColDesc
    SampleFolders     = $subfolders -join ";"
}
if (Test-Path $logFile) {
    $existingRows = @(Import-Csv $logFile)
    $matchIdx = -1
    for ($ri = 0; $ri -lt $existingRows.Count; $ri++) {
        if ($existingRows[$ri].Project -eq $projectName) { $matchIdx = $ri; break }
    }
    if ($matchIdx -ge 0) {
        $existingRows[$matchIdx] = $logRow
        $existingRows | Export-Csv $logFile -NoTypeInformation
    } else {
        $logRow | Export-Csv $logFile -Append -NoTypeInformation
    }
} else {
    $logRow | Export-Csv $logFile -NoTypeInformation
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "  Done!" -ForegroundColor Cyan
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  ID        : $projectID" -ForegroundColor Cyan
Write-Host "  PI        : $(if ($pi -eq '') { '(not specified)' } else { $pi })" -ForegroundColor White
Write-Host "  Analytics : $analyticsCol" -ForegroundColor White
Write-Host "  Desc      : $(if ($colDesc -eq '') { '(none)' } else { $colDesc })" -ForegroundColor DarkGray
Write-Host "  Trap      : $(if ($trapCol -eq '') { '(not specified)' } else { $trapCol })" -ForegroundColor White
Write-Host "  TrapDesc  : $(if ($trapColDesc -eq '') { '(none)' } else { $trapColDesc })" -ForegroundColor DarkGray
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  Projects\$(Split-Path $analyticsPath -Leaf)\" -ForegroundColor DarkGray
Write-Host "  \-- $(Split-Path $projectPath -Leaf)\" -ForegroundColor Green
if ($subfolders.Count -eq 0) {
    Write-Host "      \-- Result\" -ForegroundColor Gray
} else {
    for ($i = 0; $i -lt $subfolders.Count; $i++) {
        $isLast = ($i -eq $subfolders.Count - 1)
        $branch = if ($isLast) { "\--" } else { "+--" }
        $pipe   = if ($isLast) { "    " } else { "|   " }
        Write-Host "      $branch $($subfolders[$i])\" -ForegroundColor Green
        Write-Host "      $pipe\-- Result\" -ForegroundColor Gray
    }
}
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  Metadata  : $projectPath\project_info.json" -ForegroundColor DarkCyan
if ($colInfoChanged) { Write-Host "  ColInfo   : $colInfoFile" -ForegroundColor DarkCyan }
Write-Host "  Log       : $logFile" -ForegroundColor DarkCyan
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
