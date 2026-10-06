# Shared Windows folder / file picker helpers.
# Dot-source from a script:  . (Join-Path $PSScriptRoot "lib\Pickers.ps1")

# Projects root from config.json (fallback Z:\Proteomics\Projects)
function Get-ProjectsRoot {
    $cfgFile  = Join-Path (Split-Path $PSScriptRoot -Parent) "config.json"
    $cfg      = if (Test-Path $cfgFile) { Get-Content $cfgFile -Raw | ConvertFrom-Json } else { $null }
    $rootBase = if ($cfg -and $cfg.Root) { $cfg.Root } else { "Z:\Proteomics" }
    return (Join-Path $rootBase "Projects")
}

# Invisible topmost owner so dialogs do not open behind the console window
function New-PickerOwner {
    $owner = New-Object System.Windows.Forms.Form
    $owner.TopMost       = $true
    $owner.ShowInTaskbar = $false
    return $owner
}

# Opens the Windows folder picker.
# Returns the chosen path, "" on Cancel, or $null if the dialog cannot be shown.
function Select-Folder([string]$title, [string]$startPath) {
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description         = $title
        $dlg.ShowNewFolderButton = $false
        if ($startPath -and (Test-Path -LiteralPath $startPath)) { $dlg.SelectedPath = $startPath }
        $owner = New-PickerOwner
        $res   = $dlg.ShowDialog($owner)
        $owner.Dispose()
        if ($res -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.SelectedPath }
        return ""
    } catch {
        return $null
    }
}

# Opens the Windows file picker. $filter e.g. "CSV files (*.csv)|*.csv"
# Returns the chosen file, "" on Cancel, or $null if the dialog cannot be shown.
function Select-File([string]$title, [string]$startPath, [string]$filter) {
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title       = $title
        $dlg.Filter      = $filter
        $dlg.Multiselect = $false
        if ($startPath -and (Test-Path -LiteralPath $startPath -PathType Leaf)) {
            $dlg.InitialDirectory = Split-Path $startPath -Parent
            $dlg.FileName         = Split-Path $startPath -Leaf
        } elseif ($startPath -and (Test-Path -LiteralPath $startPath)) {
            $dlg.InitialDirectory = $startPath
        }
        $owner = New-PickerOwner
        $res   = $dlg.ShowDialog($owner)
        $owner.Dispose()
        if ($res -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.FileName }
        return ""
    } catch {
        return $null
    }
}

# Folder picker with console hint and typed fallback.
# Starts at $startPath, else the Projects root; -NoStart opens at the top of
# the tree (useful for other drives such as an external HDD).
# Returns the chosen path, or "" if the user cancelled.
function Read-FolderPath([string]$title, [string]$startPath, [switch]$NoStart) {
    if ($NoStart) {
        $startPath = ""
    } else {
        if (-not $startPath) { $startPath = Get-ProjectsRoot }
        if (-not (Test-Path -LiteralPath $startPath)) { $startPath = (Get-Location).Path }
    }
    Write-Host "  $title (window opened)..." -ForegroundColor Cyan
    $picked = Select-Folder $title $startPath
    if ($null -eq $picked) {
        $picked = (Read-Host "  $title (blank = cancel)").Trim().Trim('"')
    }
    if ($picked -ne "") { Write-Host "  Selected : $picked" -ForegroundColor White }
    return $picked
}

# File picker with console hint and typed fallback.
# Returns the chosen file, or "" if the user cancelled.
function Read-FilePath([string]$title, [string]$startPath, [string]$filter) {
    Write-Host "  $title (window opened)..." -ForegroundColor Cyan
    $picked = Select-File $title $startPath $filter
    if ($null -eq $picked) {
        $picked = (Read-Host "  $title (blank = cancel)").Trim().Trim('"')
    }
    if ($picked -ne "") { Write-Host "  Selected : $picked" -ForegroundColor White }
    return $picked
}
