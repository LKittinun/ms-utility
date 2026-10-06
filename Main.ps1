$host.UI.RawUI.WindowTitle = "Mass Spectrometry Utility Suite"
Set-Location $PSScriptRoot

$configFile  = Join-Path $PSScriptRoot "config.json"
$_cfg        = if (Test-Path $configFile) { Get-Content $configFile -Raw | ConvertFrom-Json } else { $null }
$currentRoot = if ($_cfg -and $_cfg.Root) { $_cfg.Root } else { "Z:\Proteomics" }

# Type "item" = selectable entry; "sep" = non-selectable section header
# Key = number shortcut, Hint = dim text on the right
$entries = @(
    @{ Type = "sep";  Label = "PROJECT";       Color = "Cyan" }
    @{ Type = "item"; Key = "1";  Label = "Project folder initializer"; Hint = "new or update a project";    Script = ".\01_Project_init.ps1";         Color = "White" }
    @{ Type = "item"; Key = "2";  Label = "Find project";               Hint = "search name, PI, column";    Script = ".\02_Find_project.ps1";         Color = "White" }
    @{ Type = "item"; Key = "3";  Label = "Projects overview";          Hint = "filter + Excel export";      Script = ".\03_Projects_overview.ps1";    Color = "White" }
    @{ Type = "sep";  Label = "ANALYSIS";      Color = "Cyan" }
    @{ Type = "item"; Key = "4";  Label = "Column usage report";        Hint = "runs per column";            Script = ".\04_Column_usage.ps1";         Color = "White" }
    @{ Type = "item"; Key = "5";  Label = "DIA-NN metrics";             Hint = "plots + TSV";                Script = ".\05_DIANN_metrics.ps1";        Color = "White" }
    @{ Type = "item"; Key = "6";  Label = "DIA-NN default run";         Hint = "command line";               Script = ".\06_DIANN_run.ps1";            Color = "White" }
    @{ Type = "item"; Key = "7";  Label = "Analysis report";            Hint = "Excel, per sample folder";   Script = ".\07_Report_generator.ps1";     Color = "White" }
    @{ Type = "sep";  Label = "MISCELLANEOUS"; Color = "Cyan" }
    @{ Type = "item"; Key = "8";  Label = "Bulk convert .raw to mzML";  Hint = "msConvert";                  Script = ".\08_Bulk_msConvert.ps1";       Color = "White" }
    @{ Type = "item"; Key = "9";  Label = "Contaminant check";          Hint = "mzsniffer";                  Script = ".\09_Contaminant_check.ps1";    Color = "White" }
    @{ Type = "item"; Key = "10"; Label = "Clear method files";         Hint = "*.sld  *.meth";              Script = ".\10_Clear_files.ps1";          Color = "White" }
    @{ Type = "sep";  Label = "ADMIN ONLY";    Color = "DarkYellow" }
    @{ Type = "item"; Key = "11"; Label = "Archive raw files";          Hint = "to external HDD";            Script = ".\11_Archive_raw.ps1";          Color = "Gray" }
    @{ Type = "item"; Key = "12"; Label = "Repair project order";       Hint = "password";                   Script = ".\12_Repair_project_order.ps1"; Color = "Gray" }
    @{ Type = "item"; Key = "13"; Label = "Backfill existing column";   Hint = "password";                   Script = ".\13_Backfill_column.ps1";      Color = "Gray" }
    @{ Type = "item"; Key = "14"; Label = "Sync from overview CSV";     Hint = "password";                   Script = ".\14_Sync_from_overview.ps1";   Color = "Gray" }
    @{ Type = "sep";  Label = "SETTINGS";      Color = "DarkGray" }
    @{ Type = "item"; Key = "";   Label = "Set root directory";         Hint = "";                           Script = "__SET_ROOT__";                  Color = "Gray" }
    @{ Type = "item"; Key = "";   Label = "Exit";                       Hint = "Esc";                        Script = $null;                           Color = "DarkYellow" }
)

$W      = 66     # total row width
$labelW = 30
$hintW  = $W - 2 - 2 - 2 - 2 - $labelW   # "  " + marker(2) + num(2) + "  " + label

# Indices of selectable items only
$selectable = @(0..($entries.Count - 1) | Where-Object { $entries[$_].Type -eq "item" })
$selIdx     = 0   # index into $selectable

function DrawEntry ($i) {
    $e = $entries[$i]
    [Console]::SetCursorPosition(0, $menuTop + $i)
    if ($e.Type -eq "sep") {
        $head = "  " + $e.Label + " "
        Write-Host $head -ForegroundColor $e.Color -NoNewline
        Write-Host ("-" * ($W - $head.Length)) -ForegroundColor DarkGray -NoNewline
        return
    }
    $num  = $e.Key.PadLeft(2)
    $lbl  = $e.Label.PadRight($labelW)
    $hint = $e.Hint
    if ($e.Script -eq "__SET_ROOT__") { $hint = $currentRoot }
    if ($hint.Length -gt $hintW) { $hint = "..." + $hint.Substring($hint.Length - $hintW + 3) }
    $hint = $hint.PadRight($hintW)

    Write-Host "  " -NoNewline
    if ($selectable[$selIdx] -eq $i) {
        Write-Host ("> " + $num + "  " + $lbl + $hint) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
    } else {
        $numColor = if ($e.Color -eq "White") { "Cyan" } else { "DarkGray" }
        Write-Host "  " -NoNewline
        Write-Host $num -ForegroundColor $numColor -NoNewline
        Write-Host ("  " + $lbl) -ForegroundColor $e.Color -NoNewline
        Write-Host $hint -ForegroundColor DarkGray -NoNewline
    }
}

# -- Header -------------------------------------------------------------------
Clear-Host
Write-Host ""
$banner = @(
    '   __  __ ____    _   _ _   _ _ _ _         '
    '  |  \/  / ___|  | | | | |_(_) (_) |_ _   _ '
    '  | |\/| \___ \  | | | | __| | | | __| | | |'
    '  | |  | |___) | | |_| | |_| | | | |_| |_| |'
    '  |_|  |_|____/   \___/ \__|_|_|_|\__|\__, |'
    '                                      |___/ '
)
foreach ($line in $banner) { Write-Host $line -ForegroundColor Cyan }
Write-Host "  Mass Spectrometry Utility Suite" -ForegroundColor DarkCyan -NoNewline
Write-Host ("   " + (Get-Date -Format "ddd dd MMM yyyy")) -ForegroundColor DarkGray
Write-Host ""

# -- Menu ---------------------------------------------------------------------
$menuTop = [Console]::CursorTop
for ($i = 0; $i -lt $entries.Count; $i++) {
    DrawEntry $i
    [Console]::SetCursorPosition(0, $menuTop + $i + 1)
}

# -- Footer -------------------------------------------------------------------
[Console]::SetCursorPosition(0, $menuTop + $entries.Count + 1)
Write-Host ("  " + ("-" * ($W - 2))) -ForegroundColor DarkGray
$rootOk = Test-Path -LiteralPath $currentRoot
Write-Host "  Root  " -ForegroundColor DarkGray -NoNewline
Write-Host $currentRoot -ForegroundColor White -NoNewline
if ($rootOk) { Write-Host "   [online]" -ForegroundColor Green }
else         { Write-Host "   [not found - check drive / Settings]" -ForegroundColor Red }
Write-Host "  Up/Down" -ForegroundColor Cyan -NoNewline;  Write-Host " move   " -ForegroundColor DarkGray -NoNewline
Write-Host "1-14" -ForegroundColor Cyan -NoNewline;       Write-Host " jump   " -ForegroundColor DarkGray -NoNewline
Write-Host "Enter" -ForegroundColor Cyan -NoNewline;      Write-Host " run   " -ForegroundColor DarkGray -NoNewline
Write-Host "Esc" -ForegroundColor Cyan -NoNewline;        Write-Host " exit" -ForegroundColor DarkGray
Write-Host ""
$msgRow = [Console]::CursorTop

# -- Input loop ---------------------------------------------------------------
$numBuf  = ""
$numTime = [DateTime]::MinValue
while ($true) {
    $key     = [Console]::ReadKey($true)
    $prevIdx = $selIdx

    if ($key.Key -eq [ConsoleKey]::UpArrow) {
        $selIdx = ($selIdx - 1 + $selectable.Count) % $selectable.Count
    }
    elseif ($key.Key -eq [ConsoleKey]::DownArrow) {
        $selIdx = ($selIdx + 1) % $selectable.Count
    }
    elseif ($key.Key -eq [ConsoleKey]::Home) { $selIdx = 0 }
    elseif ($key.Key -eq [ConsoleKey]::End)  { $selIdx = $selectable.Count - 1 }
    elseif ($key.KeyChar -match '^[0-9]$') {
        # Number shortcut: digits typed within 1 second combine ("1","2" -> 12)
        $now = Get-Date
        if (($now - $numTime).TotalMilliseconds -gt 1000) { $numBuf = "" }
        $numTime = $now
        $try = $numBuf + $key.KeyChar
        $hit = @(0..($selectable.Count - 1) | Where-Object { $entries[$selectable[$_]].Key -eq $try })
        if ($hit.Count -eq 0) {
            $try = "$($key.KeyChar)"
            $hit = @(0..($selectable.Count - 1) | Where-Object { $entries[$selectable[$_]].Key -eq $try })
        }
        $numBuf = $try
        if ($hit.Count -gt 0) { $selIdx = $hit[0] }
    }
    elseif ($key.Key -eq [ConsoleKey]::Enter) {
        $chosen = $entries[$selectable[$selIdx]]
        if ($null -eq $chosen.Script) {
            [Console]::SetCursorPosition(0, $msgRow)
            Write-Host "  Exiting..." -ForegroundColor DarkYellow
            return
        }
        if ($chosen.Script -eq "__SET_ROOT__") {
            [Console]::SetCursorPosition(0, $msgRow)
            . (Join-Path $PSScriptRoot "lib\Pickers.ps1")
            $newRoot = Read-FolderPath "Select the new root directory (Cancel = keep current)" $currentRoot
            if ($newRoot -ne "") { $currentRoot = $newRoot }
            [pscustomobject]@{ Root = $currentRoot } | ConvertTo-Json | Out-File $configFile -Encoding UTF8
            Clear-Host; .\Main.ps1; return
        }
        Clear-Host
        & $chosen.Script
        return
    }
    elseif ($key.Key -eq [ConsoleKey]::Escape) {
        [Console]::SetCursorPosition(0, $msgRow)
        Write-Host "  Exiting..." -ForegroundColor DarkYellow
        return
    }

    if ($prevIdx -ne $selIdx) {
        DrawEntry $selectable[$prevIdx]
        DrawEntry $selectable[$selIdx]
    }
}
