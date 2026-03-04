$host.UI.RawUI.WindowTitle = "Mass Spectrometry Utility Suite"
Set-Location $PSScriptRoot

# Type "item" = selectable entry; "sep" = non-selectable group label
$entries = @(
    @{ Type = "sep";  Label = "  Project"                                                                        }
    @{ Type = "item"; Label = "[1]  Project folder initializer";   Script = ".\01_Project_init.ps1";          Color = "White"    }
    @{ Type = "item"; Label = "[2]  Repair project order";         Script = ".\02_Repair_project_order.ps1";  Color = "White"    }
    @{ Type = "item"; Label = "[3]  Backfill existing column";     Script = ".\03_Backfill_column.ps1";       Color = "White"    }
    @{ Type = "item"; Label = "[4]  Find project";                  Script = ".\04_Find_project.ps1";         Color = "White"    }
    @{ Type = "item"; Label = "[5]  Projects overview            (Excel)"; Script = ".\05_Projects_overview.ps1";  Color = "White"    }
    @{ Type = "item"; Label = "[11] Sync from overview CSV";              Script = ".\11_Sync_from_overview.ps1"; Color = "White"    }
    @{ Type = "sep";  Label = "  Analysis"                                                                        }
    @{ Type = "item"; Label = "[6]  Column usage report";           Script = ".\06_Column_usage.ps1";          Color = "White"    }
    @{ Type = "item"; Label = "[7]  DIA-NN metrics                (plots + TSV)"; Script = ".\07_DIANN_metrics.ps1";    Color = "White" }
    @{ Type = "item"; Label = "[8]  Service report                (Excel)";       Script = ".\08_Report_generator.ps1";  Color = "White" }
    @{ Type = "sep";  Label = "  Miscellaneous"                                                                  }
    @{ Type = "item"; Label = "[9]  Bulk convert .raw to mzML     (msConvert)"; Script = ".\09_Bulk_msConvert.ps1";    Color = "DarkGray" }
    @{ Type = "item"; Label = "[10] Contaminant check             (mzsniffer)"; Script = ".\10_Contaminant_check.ps1"; Color = "DarkGray" }
    @{ Type = "item"; Label = "[99] Clear method files            (*sld *meth)"; Script = ".\99_Clear_files.ps1";     Color = "DarkGray" }
    @{ Type = "sep";  Label = ""                                                                                }
    @{ Type = "item"; Label = "Exit";                              Script = $null;                           Color = "DarkYellow" }
)

$w      = 55
$border = "=" * $w
$rule   = "-" * $w

# Indices of selectable items only
$selectable = @(0..($entries.Count - 1) | Where-Object { $entries[$_].Type -eq "item" })
$selIdx     = 0   # index into $selectable

function DrawEntry ($i) {
    $e = $entries[$i]
    [Console]::SetCursorPosition(0, $menuTop + $i)
    if ($e.Type -eq "sep") {
        $label = if ($e.Label -ne "") { "  $($e.Label) " + ("-" * [Math]::Max(0, $w - $e.Label.Length - 1)) } else { "" }
        Write-Host $label.PadRight($w + 4) -ForegroundColor DarkCyan -NoNewline
    } else {
        $highlight = ($selectable[$selIdx] -eq $i)
        $text = ("    $($e.Label)").PadRight($w + 4)
        if ($highlight) {
            Write-Host $text -ForegroundColor Black -BackgroundColor Cyan -NoNewline
        } else {
            Write-Host $text -ForegroundColor $e.Color -NoNewline
        }
    }
}

Clear-Host
Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "          Mass Spectrometry Utility Suite" -ForegroundColor Cyan
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host ""

$menuTop = [Console]::CursorTop

for ($i = 0; $i -lt $entries.Count; $i++) {
    DrawEntry $i
    [Console]::SetCursorPosition(0, $menuTop + $i + 1)
}

[Console]::SetCursorPosition(0, $menuTop + $entries.Count + 1)
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  Dir : $((Get-Location).Path)" -ForegroundColor DarkYellow
Write-Host "  Up / Down + Enter  |  Esc to exit" -ForegroundColor DarkCyan
Write-Host ""

while ($true) {
    $key     = [Console]::ReadKey($true)
    $prevIdx = $selIdx

    if ($key.Key -eq [ConsoleKey]::UpArrow) {
        $selIdx = ($selIdx - 1 + $selectable.Count) % $selectable.Count
    }
    elseif ($key.Key -eq [ConsoleKey]::DownArrow) {
        $selIdx = ($selIdx + 1) % $selectable.Count
    }
    elseif ($key.Key -eq [ConsoleKey]::Enter) {
        $chosen = $entries[$selectable[$selIdx]]
        if ($null -eq $chosen.Script) {
            [Console]::SetCursorPosition(0, $menuTop + $entries.Count + 5)
            Write-Host "  Exiting..." -ForegroundColor DarkYellow
            return
        }
        Clear-Host
        & $chosen.Script
        return
    }
    elseif ($key.Key -eq [ConsoleKey]::Escape) {
        [Console]::SetCursorPosition(0, $menuTop + $entries.Count + 5)
        Write-Host "  Exiting..." -ForegroundColor DarkYellow
        return
    }

    if ($prevIdx -ne $selIdx) {
        DrawEntry $selectable[$prevIdx]
        DrawEntry $selectable[$selIdx]
    }
}
