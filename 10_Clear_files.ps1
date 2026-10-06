$w      = 55
$border = "=" * $w
$rule   = "-" * $w

. (Join-Path $PSScriptRoot "lib\Menu.ps1")

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "   [10] Clear method files            (*sld *meth)" -ForegroundColor Cyan
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host ""

# -- Confirm ------------------------------------------------------------------
$r = Show-Menu -Items @("Run", "Back to main menu") -AllowEscape
if ($r.Action -ne "select" -or $r.Index -eq 1) { Return-ToMain; return }
Write-Host ""

. (Join-Path $PSScriptRoot "lib\Pickers.ps1")
$path = Read-FolderPath "Select the folder to clear .sld / .meth files from"
if ($path -eq "") { Return-ToMain; return }
$files = Get-ChildItem -Path $path -Recurse -Include "*.sld", "*.meth" -File

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
Write-Host ""
Write-Host "  $rule" -ForegroundColor DarkCyan
Show-NavExit
return
