$w          = 55
$border     = "=" * $w
$rule       = "-" * $w
$prohibited = @("blank", "raw_summary", "prtc", "sst", "column_usage_history", "result")

. (Join-Path $PSScriptRoot "lib\Menu.ps1")

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "   [1]  Project folder initializer" -ForegroundColor Cyan
Write-Host "        Creates project structure and logs column info" -ForegroundColor DarkCyan
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host ""

# -- Confirm ------------------------------------------------------------------
$r = Show-Menu -Items @("Run", "Back to main menu") -AllowEscape
if ($r.Action -ne "select" -or $r.Index -eq 1) { Clear-Host; .\Main.ps1; return }
Write-Host ""

# ── Root ──────────────────────────────────────────────────────────────────────
$_cfg         = if (Test-Path (Join-Path $PSScriptRoot "config.json")) { Get-Content (Join-Path $PSScriptRoot "config.json") -Raw | ConvertFrom-Json } else { $null }
$root         = if ($_cfg -and $_cfg.Root) { $_cfg.Root } else { "Z:\Proteomics" }
$projectsRoot = Join-Path $root "Projects"
Write-Host "  Root : $root" -ForegroundColor DarkGray

# ── Helpers ───────────────────────────────────────────────────────────────────
$dataDir = Join-Path $PSScriptRoot "data"
function Save-JsonList($list, [string]$file) {
    [System.IO.Directory]::CreateDirectory($dataDir) | Out-Null
    ConvertTo-Json -InputObject @($list) | Out-File $file -Encoding UTF8
}

function Write-Section([string]$title) {
    Write-Host ""
    Write-Host "  $rule" -ForegroundColor DarkCyan
    Write-Host "  $title" -ForegroundColor Cyan
    Write-Host "  $rule" -ForegroundColor DarkCyan
    Write-Host ""
}

# Spaces -> underscore, strip characters not allowed in folder names
function ConvertTo-SafeName([string]$s) {
    return (($s.Trim() -replace '[\s]+', '_') -replace '[<>:"/\\|?*]', '')
}

# Collapse whitespace + lowercase, for comparing PI names
function Get-NameKey([string]$s) {
    return ($s -replace '\s+', ' ').Trim().ToLower()
}

# Levenshtein edit distance (case-insensitive): fewest single-letter
# insert/delete/substitute edits turning $a into $b
function Get-EditDistance([string]$a, [string]$b) {
    $a = $a.ToLower(); $b = $b.ToLower(); $m = $b.Length + 1
    $d = New-Object 'int[]' (($a.Length + 1) * $m)
    for ($i = 0; $i -le $a.Length; $i++) { $d[$i * $m] = $i }
    for ($j = 0; $j -le $b.Length; $j++) { $d[$j] = $j }
    for ($i = 1; $i -le $a.Length; $i++) {
        for ($j = 1; $j -le $b.Length; $j++) {
            $cost = 1
            if ($a[$i - 1] -eq $b[$j - 1]) { $cost = 0 }
            $del = $d[($i - 1) * $m + $j] + 1
            $ins = $d[$i * $m + $j - 1] + 1
            $sub = $d[($i - 1) * $m + $j - 1] + $cost
            $d[$i * $m + $j] = [Math]::Min([Math]::Min($del, $ins), $sub)
        }
    }
    return $d[$a.Length * $m + $b.Length]
}

# Existing PI that a typed name probably means, or $null:
#  1. an existing PI appears as a whole word in the typed name
#     (e.g. "Assoc. Prof. Dr. Raphatphorn Navakanitworakul" -> "Raphatphorn")
#  2. closest existing PI within 2 edits (1 edit for names under 5 letters)
function Find-SimilarPI([string]$typed, [string[]]$pis) {
    $words = @(($typed.ToLower() -split '[^a-z]+') | Where-Object { $_ -ne "" })
    foreach ($p in $pis) {
        if ($words.Count -gt 1 -and $words -contains (Get-NameKey $p)) { return $p }
    }
    $best = $null; $bestD = [int]::MaxValue
    foreach ($p in $pis) {
        $dist = Get-EditDistance (Get-NameKey $typed) (Get-NameKey $p)
        $limit = 2
        if ([Math]::Min($typed.Length, $p.Length) -lt 5) { $limit = 1 }
        if ($dist -le $limit -and $dist -lt $bestD) { $best = $p; $bestD = $dist }
    }
    return $best
}

# ── Libraries ─────────────────────────────────────────────────────────────────
$colLibFile = Join-Path $dataDir "columns.json"
$colLib     = @()
if (Test-Path $colLibFile) {
    $loaded = Get-Content $colLibFile -Raw | ConvertFrom-Json
    if ($loaded) {
        $colLib = @($loaded)
        # Migrate old object format { ColumnID, Description } -> string array
        if ($colLib.Count -gt 0 -and $colLib[0] -is [PSCustomObject]) {
            $colLib = @($colLib | ForEach-Object { $_.Description })
            Save-JsonList $colLib $colLibFile
        }
    }
}

$trapLibFile = Join-Path $dataDir "trap_columns.json"
$trapLib     = @()
if (Test-Path $trapLibFile) {
    $tLoaded = Get-Content $trapLibFile -Raw | ConvertFrom-Json
    if ($tLoaded) { $trapLib = @($tLoaded) }
}

# All project_info.json files (Projects\<column>\<project>\), newest first.
# Used for PI and sample subfolder suggestions.
$allInfos = @()
if (Test-Path $projectsRoot) {
    $allInfos = @(
        Get-ChildItem $projectsRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            Get-ChildItem $_.FullName -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                $jp = Join-Path $_.FullName "project_info.json"
                if (Test-Path $jp) { try { Get-Content $jp -Raw | ConvertFrom-Json } catch {} }
            }
        } | Sort-Object { "$($_.Created)" } -Descending
    )
}

# ── Step functions ────────────────────────────────────────────────────────────

# Analytics column: one list of previous columns + "[+ New column ID]".
# Returns @{ ID; Desc; FromPicker }
function Select-AnalyticsColumn {
    $pairs = @()
    if (Test-Path $projectsRoot) {
        $pairs = @(
            Get-ChildItem $projectsRoot -Directory |
            Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}_' } |
            Sort-Object Name -Descending |
            ForEach-Object {
                $infoPath = Join-Path $_.FullName "column_info.json"
                $cid  = $_.Name -replace '^\d{4}-\d{2}-\d{2}_', ''
                $cdsc = ""
                if (Test-Path $infoPath) {
                    $ci = Get-Content $infoPath -Raw | ConvertFrom-Json
                    if ($ci.ColumnID)    { $cid  = $ci.ColumnID }
                    if ($ci.Description) { $cdsc = $ci.Description }
                }
                $lbl = if ($cdsc) { "$cid  [$cdsc]" } else { $cid }
                [PSCustomObject]@{ ID = $cid; Desc = $cdsc; Label = $lbl }
            }
        )
    }

    while ($true) {
        if ($pairs.Count -gt 0) {
            Write-Host "  Select column:" -ForegroundColor Cyan
            Write-Host "  (Up/Down: select   Enter: confirm)" -ForegroundColor DarkGray
            Write-Host ""
            $items = @($pairs | ForEach-Object { $_.Label }) + @("[+ New column ID]")
            $r = Show-Menu -Items $items -Special @($items.Count - 1)
            if ($r.Index -lt $pairs.Count) {
                $p = $pairs[$r.Index]
                return [PSCustomObject]@{ ID = $p.ID; Desc = $p.Desc; FromPicker = $true }
            }
            Write-Host ""
        } else {
            Write-Host "  No previous columns found." -ForegroundColor Yellow
        }

        $hint = if ($pairs.Count -gt 0) { ", blank = back to list" } else { "" }
        $id   = Read-Host "  New column ID (e.g. C20533039, no date prefix$hint)"
        $id   = ConvertTo-SafeName ($id -replace '^\s*\d{4}-\d{2}-\d{2}_', '')
        if ($id -eq "") {
            if ($pairs.Count -eq 0) { Write-Host "  Column ID cannot be empty - try again." -ForegroundColor Red }
            Write-Host ""
            continue
        }
        $match = $pairs | Where-Object { $_.ID -ieq $id } | Select-Object -First 1
        if ($match) {
            Write-Host "  Column $($match.ID) already exists - using it." -ForegroundColor Yellow
            return [PSCustomObject]@{ ID = $match.ID; Desc = $match.Desc; FromPicker = $true }
        }
        return [PSCustomObject]@{ ID = $id; Desc = ""; FromPicker = $false }
    }
}

# Column description from the library (Del removes from library). Returns string.
function Select-ColumnDescription([string]$current) {
    Write-Host ""
    Write-Host "  Column description:" -ForegroundColor Cyan
    if ($current) { Write-Host "  (current: $current)" -ForegroundColor DarkGray }
    Write-Host "  (Up/Down: select   Del: remove from library   Enter: confirm)" -ForegroundColor DarkGray
    Write-Host ""
    while ($true) {
        if ($script:colLib.Count -eq 0) {
            $nd = (Read-Host "  New description (leave blank to skip)").Trim()
            if ($nd -ne "") {
                $script:colLib = @($script:colLib) + $nd
                Save-JsonList $script:colLib $colLibFile
                Write-Host "  Saved to description library." -ForegroundColor Green
            }
            return $nd
        }
        $items = @($script:colLib) + @("[+ Add new description]")
        $sel   = 0
        for ($i = 0; $i -lt $script:colLib.Count; $i++) { if ($current -and $script:colLib[$i] -eq $current) { $sel = $i } }
        $r = Show-Menu -Items $items -Selected $sel -Special @($items.Count - 1) -AllowDelete
        if ($r.Action -eq "delete") {
            $rm = $items[$r.Index]
            $script:colLib = @($script:colLib | Where-Object { $_ -ne $rm })
            Save-JsonList $script:colLib $colLibFile
            Write-Host "  Removed '$rm' from library." -ForegroundColor DarkYellow
            continue
        }
        if ($r.Index -lt $script:colLib.Count) { return $items[$r.Index] }
        Write-Host ""
        $nd = (Read-Host "  New description (leave blank to skip)").Trim()
        if ($nd -eq "") { return "" }
        $script:colLib = @($script:colLib) + $nd
        Save-JsonList $script:colLib $colLibFile
        Write-Host "  Saved to description library." -ForegroundColor Green
    }
}

# First use date (yyyy-MM-dd), re-asks until valid. Returns string ("" = skip).
function Read-FirstUseDate {
    while ($true) {
        $d = (Read-Host "  First use date (yyyy-MM-dd, leave blank to skip)").Trim()
        if ($d -eq "") { return "" }
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParseExact($d, "yyyy-MM-dd",
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) { return $d }
        Write-Host "  Invalid format - please enter yyyy-MM-dd (e.g. 2024-03-05)" -ForegroundColor Yellow
    }
}

# New project name, re-asks on invalid input. Returns @{ Name; Path } or $null (blank = back).
function Read-NewProjectName([string]$blankHint = "back") {
    while ($true) {
        $raw = Read-Host "  Project name (blank = $blankHint)"
        if ($raw.Trim() -eq "") { return $null }
        $name = ConvertTo-SafeName $raw
        if ($name -eq "") {
            Write-Host "  Name is empty after removing invalid characters - try again." -ForegroundColor Red
            continue
        }
        if ($prohibited -contains $name.ToLower()) {
            Write-Host "  '$name' is a reserved name - try another." -ForegroundColor Red
            continue
        }
        $dup = $colProjects | Where-Object { $_.Info.Project -ieq $name } | Select-Object -First 1
        if ($dup) {
            Write-Host "  Project '$name' already exists under this column - pick it from the list instead." -ForegroundColor Red
            continue
        }
        $path = Join-Path $analyticsPath ("{0}_{1}" -f (Get-Date -Format "yyyy-MM-dd"), $name)
        if (Test-Path (Join-Path $path "project_info.json")) {
            Write-Host "  This project already exists - pick it from the list instead." -ForegroundColor Red
            continue
        }
        if (Test-Path $path) {
            Write-Host "  WARNING: project folder already exists (no project metadata - will initialize)." -ForegroundColor Yellow
        }
        return [PSCustomObject]@{ Name = $name; Path = $path }
    }
}

# PI picker: previous PIs (most recent first). Typed names that match an
# existing PI (ignoring case/spacing) reuse the existing spelling. Returns string.
function Select-PI([string]$current) {
    $seen = @{}
    $pis  = @()
    foreach ($inf in $allInfos) {
        if ($inf.PI) {
            $key = Get-NameKey $inf.PI
            if (-not $seen.ContainsKey($key)) { $seen[$key] = $true; $pis += ($inf.PI -replace '\s+', ' ').Trim() }
        }
    }
    Write-Host ""
    Write-Host "  PI:" -ForegroundColor Cyan
    if ($current) { Write-Host "  (current: $current)" -ForegroundColor DarkGray }
    Write-Host "  (Up/Down: select   Enter: confirm)" -ForegroundColor DarkGray
    Write-Host ""
    $items = @("[ Not specified ]", "[+ New PI]") + @($pis)
    $sel   = 0
    for ($i = 0; $i -lt $pis.Count; $i++) { if ($current -and (Get-NameKey $pis[$i]) -eq (Get-NameKey $current)) { $sel = $i + 2 } }
    $r = Show-Menu -Items $items -Selected $sel -Special @(0, 1)
    if ($r.Index -eq 0) { return "" }
    if ($r.Index -ge 2) { return $items[$r.Index] }
    Write-Host ""
    $new = ((Read-Host "  New PI name (blank = not specified)") -replace '\s+', ' ').Trim()
    if ($new -eq "") { return "" }
    $hit = $pis | Where-Object { (Get-NameKey $_) -eq (Get-NameKey $new) } | Select-Object -First 1
    if ($hit) {
        Write-Host "  Matches existing PI '$hit' - using that spelling." -ForegroundColor Yellow
        return $hit
    }
    $similar = Find-SimilarPI $new $pis
    if ($similar) {
        Write-Host "  Did you mean '$similar'?" -ForegroundColor Yellow
        $yr = Show-Menu -Items @("Yes - use '$similar'", "No - keep '$new' as a new PI")
        if ($yr.Index -eq 0) { return $similar }
    }
    return $new
}

# Trap column ID + description. Returns @{ ID; Desc }
function Select-TrapColumn([string]$curID, [string]$curDesc) {
    # Unique (ID, description) pairs used under this analytics column
    $seen  = @{}
    $pairs = @()
    foreach ($cp in $colProjects) {
        $tid = $cp.Info.TrapColumn
        if ($tid) {
            $tdsc = if ($cp.Info.TrapColumnDescription) { $cp.Info.TrapColumnDescription } else { "" }
            $key  = "$tid|$tdsc"
            if (-not $seen.ContainsKey($key)) {
                $seen[$key] = $true
                $lbl = if ($tdsc) { "$tid  [$tdsc]" } else { $tid }
                $pairs += [PSCustomObject]@{ ID = $tid; Desc = $tdsc; Label = $lbl }
            }
        }
    }

    Write-Host ""
    Write-Host "  Trap column:" -ForegroundColor Cyan
    if ($curID) { Write-Host "  (current: $curID$(if ($curDesc) { "  [$curDesc]" }))" -ForegroundColor DarkGray }
    Write-Host "  (Up/Down: select   Enter: confirm)" -ForegroundColor DarkGray
    Write-Host ""
    $items = @("[ No trap column ]") + @($pairs | ForEach-Object { $_.Label }) + @("[+ Type new ID]")
    $sel   = 0
    if ($curID) {
        for ($i = 0; $i -lt $pairs.Count; $i++) {
            if ($pairs[$i].ID -eq $curID -and ($sel -eq 0 -or $pairs[$i].Desc -eq $curDesc)) { $sel = $i + 1 }
        }
    }
    $r = Show-Menu -Items $items -Selected $sel -Special @(0, ($items.Count - 1))
    if ($r.Index -eq 0) { return [PSCustomObject]@{ ID = ""; Desc = "" } }
    if ($r.Index -lt $items.Count - 1) {
        $p = $pairs[$r.Index - 1]
        return [PSCustomObject]@{ ID = $p.ID; Desc = $p.Desc }
    }

    Write-Host ""
    $tid = (Read-Host "  Trap column ID (blank = no trap column)").Trim()
    if ($tid -eq "") { return [PSCustomObject]@{ ID = ""; Desc = "" } }

    # Description - scoped to descriptions already recorded for this ID
    $descs = @($colProjects | ForEach-Object { if ($_.Info.TrapColumn -eq $tid) { $_.Info.TrapColumnDescription } } |
               Where-Object { $_ } | Sort-Object -Unique)
    Write-Host ""
    Write-Host "  Trap column description:" -ForegroundColor Cyan
    Write-Host "  (Up/Down: select   Del: hide from list   Enter: confirm)" -ForegroundColor DarkGray
    Write-Host ""
    while ($true) {
        if ($descs.Count -eq 0) {
            $nd = (Read-Host "  New description (leave blank to skip)").Trim()
            if ($nd -ne "") {
                $script:trapLib = @($script:trapLib) + $nd
                Save-JsonList $script:trapLib $trapLibFile
                Write-Host "  Saved to trap column description library." -ForegroundColor Green
            }
            return [PSCustomObject]@{ ID = $tid; Desc = $nd }
        }
        $dItems = @($descs) + @("[+ Add new description]")
        $dSel   = 0
        for ($i = 0; $i -lt $descs.Count; $i++) { if ($curDesc -and $descs[$i] -eq $curDesc) { $dSel = $i } }
        $r = Show-Menu -Items $dItems -Selected $dSel -Special @($dItems.Count - 1) -AllowDelete
        if ($r.Action -eq "delete") {
            $rm    = $dItems[$r.Index]
            $descs = @($descs | Where-Object { $_ -ne $rm })
            continue
        }
        if ($r.Index -lt $descs.Count) { return [PSCustomObject]@{ ID = $tid; Desc = $dItems[$r.Index] } }
        Write-Host ""
        $nd = (Read-Host "  New description (leave blank to skip)").Trim()
        if ($nd -eq "") { return [PSCustomObject]@{ ID = $tid; Desc = "" } }
        $descs = @($descs) + $nd
        $script:trapLib = @($script:trapLib) + $nd
        Save-JsonList $script:trapLib $trapLibFile
        Write-Host "  Saved to trap column description library." -ForegroundColor Green
    }
}

# Sample subfolder checklist. $existing = folders already in the project
# (locked, always kept); $preChecked = new folders ticked so far.
# Suggestions come from subfolder names used in other projects.
# Returns string array: existing + newly ticked folders.
function Select-Subfolders([string[]]$existing, [string[]]$preChecked) {
    $names  = [System.Collections.Generic.List[string]]::new()
    $checks = [System.Collections.Generic.List[bool]]::new()
    $locked = [System.Collections.Generic.List[bool]]::new()
    $index  = @{}
    function Add-Name([string]$nm, [bool]$chk, [bool]$lck) {
        $key = $nm.ToLower()
        if ($index.ContainsKey($key)) {
            if ($chk) { $checks[$index[$key]] = $true }
            return
        }
        $index[$key] = $names.Count
        $names.Add($nm); $checks.Add($chk); $locked.Add($lck)
    }
    foreach ($nm in @($existing))   { if ($nm) { Add-Name $nm $true $true } }
    foreach ($nm in @($preChecked)) { if ($nm) { Add-Name $nm $true $false } }
    foreach ($inf in $allInfos) {
        foreach ($nm in @($inf.SampleFolders)) {
            if ($nm -and $prohibited -notcontains "$nm".ToLower()) { Add-Name "$nm" $false $false }
        }
    }

    Write-Host ""
    Write-Host "  Sample subfolders:" -ForegroundColor Cyan
    Write-Host "  (Space/Enter: tick   Up/Down: move   [ Done ]: confirm)" -ForegroundColor DarkGray
    Write-Host "  (nothing ticked = single Result\ folder in the project)" -ForegroundColor DarkGray
    Write-Host ""
    $sel = 0
    while ($true) {
        $n     = $names.Count
        $items = @($names) + @("[+ Add new subfolder]", "[ Done ]")
        $chk   = [bool[]](@($checks) + @($false, $false))
        $lck   = [bool[]](@($locked) + @($false, $false))
        $r = Show-Menu -Items $items -Selected $sel -Special @($n, ($n + 1)) -Checks $chk -Locked $lck
        for ($i = 0; $i -lt $n; $i++) { $checks[$i] = $r.Checks[$i] }
        if ($r.Index -eq $n + 1) { break }

        Write-Host ""
        $raw    = Read-Host "  New subfolder name(s), comma-separated (blank = back)"
        $parsed = @($raw -split "," | ForEach-Object { ConvertTo-SafeName $_ } | Where-Object { $_ -ne "" })
        $reservedHits = @($parsed | Where-Object { $prohibited -contains $_.ToLower() })
        if ($reservedHits.Count -gt 0) {
            Write-Host "  Reserved names not allowed as subfolders: $($reservedHits -join ', ')" -ForegroundColor Red
        }
        foreach ($nm in $parsed) {
            if ($prohibited -contains $nm.ToLower()) { continue }
            if ($index.ContainsKey($nm.ToLower()) -and $locked[$index[$nm.ToLower()]]) {
                Write-Host "  '$nm' already exists in this project." -ForegroundColor Yellow
                continue
            }
            Add-Name $nm $true $false
        }
        Write-Host ""
        $sel = $names.Count + 1   # land on [ Done ]
    }

    $result = @(for ($i = 0; $i -lt $names.Count; $i++) { if ($checks[$i]) { $names[$i] } })
    return ,$result
}

# ── Analytics column ──────────────────────────────────────────────────────────
Write-Section "Analytics column"
$col               = Select-AnalyticsColumn
$analyticsCol      = $col.ID
$colDesc           = $col.Desc
if (-not $col.FromPicker) { $colDesc = Select-ColumnDescription "" }

# Resolve analytics column folder (date-prefixed, e.g. 2026-03-02_C20533039)
$analyticsPath = $null
if (Test-Path $projectsRoot) {
    $existingColDir = Get-ChildItem $projectsRoot -Directory |
        Where-Object { $_.Name -like "*_$analyticsCol" } |
        Sort-Object Name |
        Select-Object -First 1
    if ($existingColDir) { $analyticsPath = $existingColDir.FullName }
}
if (-not $analyticsPath) { $analyticsPath = Join-Path $projectsRoot ("{0}_{1}" -f (Get-Date -Format "yyyy-MM-dd"), $analyticsCol) }
$logFile = Join-Path $analyticsPath "column_log.csv"

# Column info JSON
$colInfoFile    = Join-Path $analyticsPath "column_info.json"
$colInfoData    = $null
$colInfoChanged = $false
if (Test-Path $colInfoFile) { $colInfoData = Get-Content $colInfoFile -Raw | ConvertFrom-Json }

if ($null -eq $colInfoData) {
    Write-Section "New column detected - enter column details (all optional):"
    $colFirstUse    = Read-FirstUseDate
    $colInfoChanged = $true
} else {
    $colFirstUse = if ($colInfoData.FirstUseDate) { $colInfoData.FirstUseDate } else { "" }
    Write-Host ""
    Write-Host "  Column: $analyticsCol" -ForegroundColor Cyan
}

# Projects already under this column, newest first
$colProjects = @()
if (Test-Path $analyticsPath) {
    $colProjects = @(
        Get-ChildItem $analyticsPath -Directory |
        Where-Object { Test-Path (Join-Path $_.FullName "project_info.json") } |
        Sort-Object Name -Descending |
        ForEach-Object {
            [PSCustomObject]@{ Dir = $_; Info = (Get-Content (Join-Path $_.FullName "project_info.json") -Raw | ConvertFrom-Json) }
        }
    )
}

# ── Project selection ─────────────────────────────────────────────────────────
Write-Section "Project"
$existingInfo = $null
$projectPath  = $null
$projectName  = ""
while ($true) {
    Write-Host "  (Up/Down: select   Enter: confirm)" -ForegroundColor DarkGray
    Write-Host ""
    $items = @("[ New project ]") + @($colProjects | ForEach-Object {
        $lbl = $_.Info.Project
        if ($_.Info.PI) { $lbl += "  ($($_.Info.PI))" }
        $lbl
    })
    $r = Show-Menu -Items $items -Special @(0)
    if ($r.Index -gt 0) {
        $chosen       = $colProjects[$r.Index - 1]
        $projectPath  = $chosen.Dir.FullName
        $existingInfo = $chosen.Info
        $projectName  = $existingInfo.Project
        Write-Host "  Existing project: $projectName" -ForegroundColor Yellow
        Write-Host "  (fields are kept unless you change them)" -ForegroundColor DarkGray
        break
    }
    Write-Host ""
    $np = Read-NewProjectName "back to list"
    if ($np) { $projectName = $np.Name; $projectPath = $np.Path; break }
    Write-Host ""
}

# ── PI ────────────────────────────────────────────────────────────────────────
if ($existingInfo) {
    $pi = if ($existingInfo.PI) { $existingInfo.PI } else { "" }
    Write-Host ""
    Write-Host "  PI: $(if ($pi) { $pi } else { '(not specified)' })" -ForegroundColor Cyan
} else {
    $pi = Select-PI ""
}

# ── Trap column ───────────────────────────────────────────────────────────────
$curTrap     = if ($existingInfo -and $existingInfo.TrapColumn)            { $existingInfo.TrapColumn }            else { "" }
$curTrapDesc = if ($existingInfo -and $existingInfo.TrapColumnDescription) { $existingInfo.TrapColumnDescription } else { "" }
$trap        = Select-TrapColumn $curTrap $curTrapDesc
$trapCol     = $trap.ID
$trapColDesc = $trap.Desc

# ── Sample subfolders ─────────────────────────────────────────────────────────
$existingFolders = if ($existingInfo -and $existingInfo.SampleFolders) { @($existingInfo.SampleFolders) } else { @() }
$subfolders      = Select-Subfolders $existingFolders @()

# ── Project number + ID ───────────────────────────────────────────────────────
if ($existingInfo) {
    $projectID = $existingInfo.ProjectID
    $projectNo = $existingInfo.ProjectNo
} else {
    $projectNo = $colProjects.Count + 1
    $projectID = -join ((65..90) + (48..57) | Get-Random -Count 8 | ForEach-Object { [char]$_ })
}

# ── Preview (tree) + edit loop ────────────────────────────────────────────────
while ($true) {
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

    $r = Show-Menu -Items @("Yes, create folders", "Edit a field", "No, cancel") -AllowEscape
    if ($r.Action -eq "cancel" -or $r.Index -eq 2) {
        Write-Host "  Cancelled." -ForegroundColor DarkYellow
        Write-Host "  $rule" -ForegroundColor DarkCyan
        Show-NavExit
        return
    }
    if ($r.Index -eq 0) { break }

    # -- Edit a field --
    $fields = @()
    if (-not $existingInfo) { $fields += "Project name" }
    $fields += @("PI", "Trap column", "Sample subfolders")
    if ($colInfoChanged)    { $fields += @("Column description", "Column first use date") }
    $fields += "[ Back to preview ]"
    Write-Host ""
    Write-Host "  Which field?" -ForegroundColor Cyan
    Write-Host "  (analytics column cannot be changed here - cancel and start again)" -ForegroundColor DarkGray
    Write-Host ""
    $fr = Show-Menu -Items $fields -Special @($fields.Count - 1) -AllowEscape
    if ($fr.Action -eq "cancel") { continue }
    switch ($fields[$fr.Index]) {
        "Project name" {
            Write-Host ""
            $np = Read-NewProjectName "keep '$projectName'"
            if ($np) { $projectName = $np.Name; $projectPath = $np.Path }
        }
        "PI" { $pi = Select-PI $pi }
        "Trap column" {
            $trap        = Select-TrapColumn $trapCol $trapColDesc
            $trapCol     = $trap.ID
            $trapColDesc = $trap.Desc
        }
        "Sample subfolders" {
            $newOnes    = @($subfolders | Where-Object { $existingFolders -inotcontains $_ })
            $subfolders = Select-Subfolders $existingFolders $newOnes
        }
        "Column description"    { $colDesc = Select-ColumnDescription $colDesc }
        "Column first use date" { Write-Host ""; $colFirstUse = Read-FirstUseDate }
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
    Show-NavExit
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
    PI                = if ($pi -eq "")       { $null } else { $pi }
    AnalyticsColumn   = $analyticsCol
    ColumnDescription = if ($colDesc -eq "")  { $null } else { $colDesc }
    TrapColumn            = if ($trapCol -eq "")     { $null } else { $trapCol }
    TrapColumnDescription = if ($trapColDesc -eq "") { $null } else { $trapColDesc }
    SampleFolders     = $subfolders -join ";"
}
if (Test-Path $logFile) {
    $existingRows = @(Import-Csv $logFile)
    $matchIdx = -1
    for ($ri = 0; $ri -lt $existingRows.Count; $ri++) {
        if ($existingRows[$ri].ProjectID -and $existingRows[$ri].ProjectID -eq $projectID) { $matchIdx = $ri; break }
    }
    if ($matchIdx -lt 0) {
        for ($ri = 0; $ri -lt $existingRows.Count; $ri++) {
            if ($existingRows[$ri].Project -eq $projectName) { $matchIdx = $ri; break }
        }
    }
    $colOrder = @("ProjectID","ProjectNo","Date","Project","PI","AnalyticsColumn","ColumnDescription","TrapColumn","TrapColumnDescription","SampleFolders")
    if ($matchIdx -ge 0) {
        $existingRows[$matchIdx] = $logRow
        $existingRows | Select-Object $colOrder | Export-Csv $logFile -NoTypeInformation -Encoding UTF8
    } else {
        $logRow | Export-Csv $logFile -Append -NoTypeInformation -Encoding UTF8
    }
} else {
    $logRow | Export-Csv $logFile -NoTypeInformation -Encoding UTF8
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
Write-Host ""
Write-Host "  $rule" -ForegroundColor DarkCyan
Show-NavExit
