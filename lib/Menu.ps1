# Shared arrow-key menu helpers.
# Dot-source from a script:  . (Join-Path $PSScriptRoot "lib\Menu.ps1")

# Scrolling arrow-key menu drawn at the current cursor position.
#   -Special     indexes drawn in DarkYellow (e.g. "[+ Add new]"); never ticked/deleted
#   -AllowDelete Del on a normal item returns Action "delete"
#   -AllowEscape Esc returns Action "cancel"
#   -Checks      bool per item -> checklist mode: Space/Enter ticks a normal item,
#                Enter on a Special item returns "select"; final ticks in .Checks
#   -Locked      bool per item; locked items stay ticked and cannot be changed
# Long lists scroll (PgUp/PgDn/Home/End also work).
# Returns [pscustomobject]@{ Action = "select"|"delete"|"cancel"; Index; Checks }
function Show-Menu {
    param(
        [string[]]$Items,
        [int]$Selected = 0,
        [int[]]$Special = @(),
        [switch]$AllowDelete,
        [switch]$AllowEscape,
        [bool[]]$Checks = $null,
        [bool[]]$Locked = $null,
        [int]$MaxRows = 15
    )
    $n = $Items.Count
    if ($n -eq 0) { return [pscustomobject]@{ Action = "cancel"; Index = -1; Checks = $Checks } }

    $winH = 30; $winW = 80
    try { $winH = [Console]::WindowHeight; $winW = [Console]::WindowWidth } catch {}
    $rowCap = [Math]::Max(5, [Math]::Min($MaxRows, $winH - 10))
    $rows   = [Math]::Min($n, $rowCap)
    $more   = $n -gt $rows
    $lines  = $rows
    if ($more) { $lines = $rows + 1 }

    $labels = @(for ($i = 0; $i -lt $n; $i++) {
        $t = $Items[$i]
        if ($null -ne $Checks -and $Special -notcontains $i) {
            $box = "[ ] "
            if ($Checks[$i]) { $box = "[x] " }
            $t = $box + $t
            if ($null -ne $Locked -and $Locked[$i]) { $t += "  (existing)" }
        }
        "    " + $t
    })
    $longest = ($labels | Measure-Object -Property Length -Maximum).Maximum
    $lineW   = [Math]::Max(20, [Math]::Min($winW - 1, [Math]::Max(59, $longest)))

    # Reserve the lines first so a scrolling console buffer does not shift the menu
    for ($i = 0; $i -lt $lines; $i++) { Write-Host "" }
    $top    = [Console]::CursorTop - $lines
    $sel    = [Math]::Max(0, [Math]::Min($Selected, $n - 1))
    $offset = 0

    while ($true) {
        if ($sel -lt $offset)         { $offset = $sel }
        if ($sel -ge $offset + $rows) { $offset = $sel - $rows + 1 }
        for ($r = 0; $r -lt $rows; $r++) {
            $i = $offset + $r
            $t = $labels[$i]
            if ($null -ne $Checks -and $Special -notcontains $i) {
                $box = "[ ] "
                if ($Checks[$i]) { $box = "[x] " }
                $t = "    " + $box + $t.Substring(8)
            }
            if ($t.Length -gt $lineW) { $t = $t.Substring(0, $lineW - 3) + "..." }
            $t = $t.PadRight($lineW)
            [Console]::SetCursorPosition(0, $top + $r)
            if ($i -eq $sel)                                 { Write-Host $t -ForegroundColor Black -BackgroundColor Cyan -NoNewline }
            elseif ($Special -contains $i)                   { Write-Host $t -ForegroundColor DarkYellow -NoNewline }
            elseif ($null -ne $Locked -and $Locked[$i])      { Write-Host $t -ForegroundColor DarkGray -NoNewline }
            else                                             { Write-Host $t -ForegroundColor White -NoNewline }
        }
        if ($more) {
            [Console]::SetCursorPosition(0, $top + $rows)
            $info = "    ($($sel + 1) of $n - Up/Down or PgUp/PgDn to scroll)"
            Write-Host $info.PadRight($lineW) -ForegroundColor DarkGray -NoNewline
        }

        $k        = [Console]::ReadKey($true)
        $isNormal = $Special -notcontains $sel
        $canTick  = $null -ne $Checks -and $isNormal -and -not ($null -ne $Locked -and $Locked[$sel])
        $result   = $null
        if     ($k.Key -eq [ConsoleKey]::UpArrow)   { $sel = ($sel - 1 + $n) % $n }
        elseif ($k.Key -eq [ConsoleKey]::DownArrow) { $sel = ($sel + 1) % $n }
        elseif ($k.Key -eq [ConsoleKey]::PageUp)    { $sel = [Math]::Max(0, $sel - $rows) }
        elseif ($k.Key -eq [ConsoleKey]::PageDown)  { $sel = [Math]::Min($n - 1, $sel + $rows) }
        elseif ($k.Key -eq [ConsoleKey]::Home)      { $sel = 0 }
        elseif ($k.Key -eq [ConsoleKey]::End)       { $sel = $n - 1 }
        elseif ($k.Key -eq [ConsoleKey]::Spacebar)  { if ($canTick) { $Checks[$sel] = -not $Checks[$sel] } }
        elseif ($k.Key -eq [ConsoleKey]::Enter) {
            if ($null -ne $Checks -and $isNormal) { if ($canTick) { $Checks[$sel] = -not $Checks[$sel] } }
            else { $result = "select" }
        }
        elseif ($k.Key -eq [ConsoleKey]::Delete -and $AllowDelete -and $isNormal) { $result = "delete" }
        elseif ($k.Key -eq [ConsoleKey]::Escape -and $AllowEscape)                { $result = "cancel" }

        if ($result) {
            [Console]::SetCursorPosition(0, $top + $lines)
            return [pscustomobject]@{ Action = $result; Index = $sel; Checks = $Checks }
        }
    }
}

# -- Navigation ---------------------------------------------------------------
# Main.ps1 runs scripts in a loop and sets $global:MsInMain. A script just
# `return`s to get back to the menu; Main redraws itself. When a script is run
# on its own (not from Main), Return-ToMain launches Main instead.

# Go back to the main menu. Caller should `return` right after.
function Return-ToMain {
    Clear-Host
    if (-not $global:MsInMain) { & (Join-Path (Split-Path $PSScriptRoot -Parent) "Main.ps1") }
}

# Leave the whole suite (Main stops after the current script returns).
# Caller should `return` right after.
function Request-Exit {
    $global:MsExit = $true
    Write-Host "  Exiting..." -ForegroundColor DarkYellow
}

# "Back to main menu / Exit" footer. Caller should `return` right after.
function Show-NavExit {
    Write-Host ""
    $r = Show-Menu -Items @("Back to main menu", "Exit") -AllowEscape
    if ($r.Action -eq "select" -and $r.Index -eq 0) { Return-ToMain }
    else { Request-Exit }
}
