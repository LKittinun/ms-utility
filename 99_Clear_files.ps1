$w      = 55
$border = "=" * $w
$rule   = "-" * $w

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "  [99]  Clear method files            (*sld *meth)" -ForegroundColor Cyan
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

$path = Read-Host "Insert directory, leave blank for a current location"
if ($path -eq "") { $path = (Get-Location).Path }
$files = Get-ChildItem -Path "$path\*" -Recurse -Include *.sld, *.meth -File

if ($files.Count -eq 0) {
    Write-Host "  No *sld and *meth files found" -ForegroundColor Yellow
} else {
    Write-Host "  These files will be removed:" -ForegroundColor Yellow
    $files | ForEach-Object { Write-Host "    $($_.FullName)" }
    Write-Host ""
    $confirm = Read-Host "  Confirm removal? y = yes"
    if ($confirm -eq "y") {
        Remove-Item $files
        Write-Host "  All files removed." -ForegroundColor Green
    } else {
        Write-Host "  Cancelled." -ForegroundColor DarkYellow
    }
}

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
