$w      = 55
$border = "=" * $w
$rule   = "-" * $w

Write-Host ""
Write-Host "  $border" -ForegroundColor DarkCyan
Write-Host "   [6]  DIA-NN default run              (command line)" -ForegroundColor Cyan
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

# -- Load / auto-create diann_config.json -------------------------------------
$diannConfigPath = Join-Path $PSScriptRoot "diann_config.json"
if (-not (Test-Path $diannConfigPath)) {
    $default = [pscustomobject]@{
        DiannExe    = "C:\DIA-NN\2.6.0\diann.exe"
        Library     = "C:\DIA-NN\Library\HS_400_1000_145_1450.predicted.speclib"
        MainFasta   = "C:\DIA-NN\Library\Homo_sapiens_reviewed_2023_12_10.fasta"
        ContamFasta = "C:\DIA-NN\2.6.0\camprotR_240512_cRAP_20190401_full_tags.fasta"
        Threads     = 12
        MassAcc     = 15
        MassAccMs1  = 7
    }
    $default | ConvertTo-Json | Out-File $diannConfigPath -Encoding UTF8
    Write-Host "  diann_config.json created at: $diannConfigPath" -ForegroundColor Yellow
    Write-Host "  Edit this file to change DIA-NN paths or default parameter values." -ForegroundColor DarkYellow
    Write-Host ""
}
$diannCfg = Get-Content $diannConfigPath -Raw | ConvertFrom-Json

# -- Validate diann.exe -------------------------------------------------------
if (-not (Test-Path $diannCfg.DiannExe)) {
    Write-Host "  WARNING: diann.exe not found at: $($diannCfg.DiannExe)" -ForegroundColor Yellow
    Write-Host "           Continuing in case 'diann' is on PATH." -ForegroundColor DarkYellow
    Write-Host ""
}

# -- Path input ---------------------------------------------------------------
$path = Read-Host "  Project directory (blank = current location)"
if ($path -eq "") { $path = (Get-Location).Path }

if (-not (Test-Path $path)) {
    Write-Host "  Path not found: $path" -ForegroundColor Red
    $targetFolder = $null
} else {
    # -- Detect candidate folders with .raw files (direct children only) ----------
    $candidates = @()

    $rootRaws = @(Get-ChildItem -Path $path -Filter *.raw -File -ErrorAction SilentlyContinue)
    if ($rootRaws.Count -gt 0) {
        $candidates += [pscustomobject]@{ FullName = $path; Name = "(this folder)"; RawCount = $rootRaws.Count }
    }

    $subDirs = @(Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue)
    foreach ($d in $subDirs) {
        $r = @(Get-ChildItem -Path $d.FullName -Filter *.raw -File -ErrorAction SilentlyContinue)
        if ($r.Count -gt 0) {
            $candidates += [pscustomobject]@{ FullName = $d.FullName; Name = $d.Name; RawCount = $r.Count }
        }
    }

    $targetFolder = $null

    if ($candidates.Count -eq 0) {
        Write-Host ""
        Write-Host "  No .raw files found in $path or its immediate subfolders." -ForegroundColor Red
    } elseif ($candidates.Count -eq 1) {
        $targetFolder = $candidates[0].FullName
        Write-Host ""
        Write-Host "  Using: $($candidates[0].Name)  [$($candidates[0].RawCount) raw files]" -ForegroundColor Cyan
    } else {
        Write-Host ""
        Write-Host "  Select subfolder to run DIA-NN on:" -ForegroundColor Cyan
        $pSel = 0
        $pTop = [Console]::CursorTop
        for ($pi = 0; $pi -lt $candidates.Count; $pi++) {
            $label = "  $($candidates[$pi].Name)  [$($candidates[$pi].RawCount) raw files]"
            [Console]::SetCursorPosition(0, $pTop + $pi)
            if ($pi -eq 0) {
                Write-Host ("> $label").PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
            } else {
                Write-Host ("  $label").PadRight($w + 4) -ForegroundColor White -NoNewline
            }
        }
        [Console]::SetCursorPosition(0, $pTop + $candidates.Count)
        :pickerLoop while ($true) {
            $pk = [Console]::ReadKey($true)
            if ($pk.Key -eq [ConsoleKey]::UpArrow -or $pk.Key -eq [ConsoleKey]::DownArrow) {
                $oldSel = $pSel
                if ($pk.Key -eq [ConsoleKey]::UpArrow) { $pSel = ($pSel - 1 + $candidates.Count) % $candidates.Count }
                else { $pSel = ($pSel + 1) % $candidates.Count }
                [Console]::SetCursorPosition(0, $pTop + $oldSel)
                Write-Host ("    $($candidates[$oldSel].Name)  [$($candidates[$oldSel].RawCount) raw files]").PadRight($w + 4) -ForegroundColor White -NoNewline
                [Console]::SetCursorPosition(0, $pTop + $pSel)
                Write-Host (">   $($candidates[$pSel].Name)  [$($candidates[$pSel].RawCount) raw files]").PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
            } elseif ($pk.Key -eq [ConsoleKey]::Enter) {
                $targetFolder = $candidates[$pSel].FullName
                break pickerLoop
            } elseif ($pk.Key -eq [ConsoleKey]::Escape) {
                Clear-Host; .\Main.ps1; return
            }
        }
        Write-Host ""
    }
}

if ($null -eq $targetFolder) {
    # -- Navigation (error / no files case) -----------------------------------
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
}

# -- Parameter defaults (set once, preserved across re-runs) ------------------
$massAcc    = [int]$diannCfg.MassAcc
$massAccMs1 = [int]$diannCfg.MassAccMs1
$threads    = [int]$diannCfg.Threads
$mbr        = $true
$sortByDate = $false

# -- Run loop (Re-run re-uses $targetFolder without re-asking the folder) -----
:runLoop while ($true) {

# -- Collect .raw files -------------------------------------------------------
$rawFiles = @(Get-ChildItem -Path $targetFolder -Filter *.raw -File -ErrorAction SilentlyContinue | Sort-Object Name)

if ($rawFiles.Count -eq 0) {
    Write-Host ""
    Write-Host "  No .raw files found in: $(Split-Path $targetFolder -Leaf)" -ForegroundColor Red
    $nItems = @("Re-run same folder", "Back to main menu", "Exit")
    $nSel   = 0
    Write-Host ""
    Write-Host "  $rule" -ForegroundColor DarkCyan
    $nTop = [Console]::CursorTop
    for ($ni = 0; $ni -lt $nItems.Count; $ni++) {
        [Console]::SetCursorPosition(0, $nTop + $ni)
        $color = if ($ni -eq 2) { "DarkYellow" } else { "Cyan" }
        if ($ni -eq 0) {
            Write-Host ("  > " + $nItems[$ni]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
        } else {
            Write-Host ("    " + $nItems[$ni]).PadRight($w + 4) -ForegroundColor $color -NoNewline
        }
    }
    [Console]::SetCursorPosition(0, $nTop + $nItems.Count)
    while ($true) {
        $k = [Console]::ReadKey($true)
        if ($k.Key -eq [ConsoleKey]::UpArrow -or $k.Key -eq [ConsoleKey]::DownArrow) {
            $p = $nSel
            if ($k.Key -eq [ConsoleKey]::UpArrow) { $nSel = ($nSel - 1 + $nItems.Count) % $nItems.Count }
            else { $nSel = ($nSel + 1) % $nItems.Count }
            $pc = if ($p -eq 2) { "DarkYellow" } else { "Cyan" }
            [Console]::SetCursorPosition(0, $nTop + $p);    Write-Host ("    " + $nItems[$p]).PadRight($w + 4)    -ForegroundColor $pc -NoNewline
            [Console]::SetCursorPosition(0, $nTop + $nSel); Write-Host ("  > " + $nItems[$nSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
        } elseif ($k.Key -eq [ConsoleKey]::Enter) {
            if ($nSel -eq 0) { continue runLoop }
            if ($nSel -eq 1) { Clear-Host; .\Main.ps1; return }
            [Console]::SetCursorPosition(0, $nTop + $nItems.Count + 1); Write-Host "  Exiting..." -ForegroundColor DarkYellow; return
        } elseif ($k.Key -eq [ConsoleKey]::Escape) {
            Clear-Host; .\Main.ps1; return
        }
    }
}

# -- Check for locked files (still being acquired) ----------------------------
$lockedFiles = @($rawFiles | Where-Object {
    try {
        $fs = [System.IO.File]::Open($_.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
        $fs.Close(); $fs.Dispose(); $false
    } catch { $true }
})

Write-Host ""
Write-Host "  $rule" -ForegroundColor DarkCyan
Write-Host "  Folder : $(Split-Path $targetFolder -Leaf)" -ForegroundColor Cyan
Write-Host "  Raw files ($($rawFiles.Count)):" -ForegroundColor White
$rawFiles | ForEach-Object {
    $isLocked = $lockedFiles.FullName -contains $_.FullName
    if ($isLocked) {
        Write-Host "    $($_.Name)  [LOCKED - still acquiring]" -ForegroundColor Yellow
    } else {
        Write-Host "    $($_.Name)" -ForegroundColor DarkGray
    }
}
Write-Host ""

$abortRun = $false
if ($lockedFiles.Count -gt 0) {
    Write-Host "  WARNING: $($lockedFiles.Count) file(s) are still locked (acquisition in progress)." -ForegroundColor Yellow
    $skipAnswer = Read-Host "  Skip locked files and continue with the rest? [Y/N]"
    if ($skipAnswer -match '^[Yy]') {
        $rawFiles = @($rawFiles | Where-Object { $lockedFiles.FullName -notcontains $_.FullName })
        Write-Host "  Continuing with $($rawFiles.Count) file(s)." -ForegroundColor Cyan
        Write-Host ""
        if ($rawFiles.Count -eq 0) {
            Write-Host "  No files remaining to process." -ForegroundColor Red
            $abortRun = $true
        }
    } else {
        $abortRun = $true
    }
}

if ($abortRun) {
    $nItems = @("Re-run same folder", "Back to main menu", "Exit")
    $nSel   = 0
    Write-Host ""
    Write-Host "  $rule" -ForegroundColor DarkCyan
    $nTop = [Console]::CursorTop
    for ($ni = 0; $ni -lt $nItems.Count; $ni++) {
        [Console]::SetCursorPosition(0, $nTop + $ni)
        $color = if ($ni -eq 2) { "DarkYellow" } else { "Cyan" }
        if ($ni -eq 0) {
            Write-Host ("  > " + $nItems[$ni]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
        } else {
            Write-Host ("    " + $nItems[$ni]).PadRight($w + 4) -ForegroundColor $color -NoNewline
        }
    }
    [Console]::SetCursorPosition(0, $nTop + $nItems.Count)
    while ($true) {
        $k = [Console]::ReadKey($true)
        if ($k.Key -eq [ConsoleKey]::UpArrow -or $k.Key -eq [ConsoleKey]::DownArrow) {
            $p = $nSel
            if ($k.Key -eq [ConsoleKey]::UpArrow) { $nSel = ($nSel - 1 + $nItems.Count) % $nItems.Count }
            else { $nSel = ($nSel + 1) % $nItems.Count }
            $pc = if ($p -eq 2) { "DarkYellow" } else { "Cyan" }
            [Console]::SetCursorPosition(0, $nTop + $p);    Write-Host ("    " + $nItems[$p]).PadRight($w + 4)    -ForegroundColor $pc -NoNewline
            [Console]::SetCursorPosition(0, $nTop + $nSel); Write-Host ("  > " + $nItems[$nSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
        } elseif ($k.Key -eq [ConsoleKey]::Enter) {
            if ($nSel -eq 0) { continue runLoop }
            if ($nSel -eq 1) { Clear-Host; .\Main.ps1; return }
            [Console]::SetCursorPosition(0, $nTop + $nItems.Count + 1); Write-Host "  Exiting..." -ForegroundColor DarkYellow; return
        } elseif ($k.Key -eq [ConsoleKey]::Escape) {
            Clear-Host; .\Main.ps1; return
        }
    }
}

# -- Check for existing .quant files ------------------------------------------
$quantFiles = @($rawFiles | Where-Object { Test-Path ($_.FullName + ".quant") })
$useQuant   = $false
if ($quantFiles.Count -gt 0) {
    Write-Host "  Found $($quantFiles.Count) existing .quant file(s) in this folder." -ForegroundColor Yellow
    $qAnswer = Read-Host "  Reuse them? (skips first-pass quantification) [Y/N]"
    $useQuant = ($qAnswer -match '^[Yy]')
    Write-Host ""
}

# -- Output paths -------------------------------------------------------------
$resultDir     = Join-Path $targetFolder "Result"
$outParquet    = Join-Path $resultDir "report.parquet"
$outLibParquet = Join-Path $resultDir "report-lib.parquet"

# -- Preview / confirm loop ---------------------------------------------------
:previewLoop while ($true) {
    $orderedFiles = if ($sortByDate) { @($rawFiles | Sort-Object LastWriteTime) } else { @($rawFiles | Sort-Object Name) }
    $diannArgs = @()
    foreach ($f in $orderedFiles) { $diannArgs += @("--f", $f.FullName) }
    $diannArgs += @(
        "--lib",                $diannCfg.Library,
        "--threads",            "$threads",
        "--verbose",            "1",
        "--out",                $outParquet,
        "--qvalue",             "0.01",
        "--matrices",
        "--out-lib",            $outLibParquet,
        "--gen-spec-lib",
        "--xic",
        "--fasta",              $diannCfg.ContamFasta,
        "--cont-quant-exclude", "cRAP-",
        "--fasta",              $diannCfg.MainFasta,
        "--met-excision",
        "--min-pep-len",        "7",
        "--max-pep-len",        "30",
        "--min-pr-mz",          "400",
        "--max-pr-mz",          "1000",
        "--min-pr-charge",      "1",
        "--max-pr-charge",      "4",
        "--min-fr-mz",          "145",
        "--max-fr-mz",          "1450",
        "--cut",                "K*,R*",
        "--missed-cleavages",   "1",
        "--unimod4",
        "--mass-acc",           "$massAcc",
        "--mass-acc-ms1",       "$massAccMs1"
    )
    if ($mbr)      { $diannArgs += @("--reanalyse", "--rt-profiling") }
    if ($useQuant) { $diannArgs += "--use-quant" }

    $mbrText = if ($mbr) { "yes" } else { "no" }
    $mbrPart = if ($mbr) { "--reanalyse --rt-profiling" } else { "(MBR disabled)" }
    $quPart  = if ($useQuant) { " --use-quant" } else { "" }
    Write-Host ""
    Write-Host "  $rule" -ForegroundColor DarkCyan
    Write-Host "  Command preview:" -ForegroundColor Cyan
    Write-Host "    $($diannCfg.DiannExe)" -ForegroundColor White
    $sortText = if ($sortByDate) { "acquisition date" } else { "name" }
    Write-Host "    --f  [$($rawFiles.Count) raw file(s), sorted by $sortText]" -ForegroundColor DarkGray
    Write-Host "    --lib $($diannCfg.Library)" -ForegroundColor DarkGray
    Write-Host "    --out $outParquet" -ForegroundColor DarkGray
    Write-Host "    --threads $threads  --mass-acc $massAcc  --mass-acc-ms1 $massAccMs1" -ForegroundColor DarkGray
    Write-Host "    MBR: $mbrText  $mbrPart$quPart  (plus standard flags)" -ForegroundColor DarkGray
    Write-Host "  $rule" -ForegroundColor DarkCyan
    Write-Host ""

    $pItems = @("Run DIA-NN", "Edit parameters", "Back to main menu")
    $pSel   = 0
    $pTop   = [Console]::CursorTop
    for ($pi = 0; $pi -lt $pItems.Count; $pi++) {
        [Console]::SetCursorPosition(0, $pTop + $pi)
        if ($pi -eq 0) {
            Write-Host ("  > " + $pItems[$pi]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
        } else {
            Write-Host ("    " + $pItems[$pi]).PadRight($w + 4) -ForegroundColor DarkCyan -NoNewline
        }
    }
    [Console]::SetCursorPosition(0, $pTop + $pItems.Count)

    :pickerLoop while ($true) {
        $pk = [Console]::ReadKey($true)
        if ($pk.Key -eq [ConsoleKey]::UpArrow -or $pk.Key -eq [ConsoleKey]::DownArrow) {
            $pp = $pSel
            if ($pk.Key -eq [ConsoleKey]::UpArrow) { $pSel = ($pSel - 1 + $pItems.Count) % $pItems.Count }
            else { $pSel = ($pSel + 1) % $pItems.Count }
            [Console]::SetCursorPosition(0, $pTop + $pp);   Write-Host ("    " + $pItems[$pp]).PadRight($w + 4)   -ForegroundColor DarkCyan -NoNewline
            [Console]::SetCursorPosition(0, $pTop + $pSel); Write-Host ("  > " + $pItems[$pSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
        } elseif ($pk.Key -eq [ConsoleKey]::Enter) {
            if ($pSel -eq 0) { break previewLoop }
            if ($pSel -eq 1) {
                Write-Host ""
                $inMa  = Read-Host "  Mass accuracy MS2 ppm    (current: $massAcc)"
                $inMs1 = Read-Host "  Mass accuracy MS1 ppm    (current: $massAccMs1)"
                $inThr = Read-Host "  Threads                  (current: $threads)"
                $mbrCur  = if ($mbr)        { "Y" } else { "N" }
                $sortCur = if ($sortByDate) { "D" } else { "N" }
                $inMbr  = Read-Host "  Match between runs (MBR) (current: $mbrCur) [Y/N]"
                $inSort = Read-Host "  File order               (current: $sortCur) [N]ame / [D]ate acquired"
                if ($inMa   -ne "") { $massAcc    = [int]$inMa  }
                if ($inMs1  -ne "") { $massAccMs1 = [int]$inMs1 }
                if ($inThr  -ne "") { $threads    = [int]$inThr }
                if ($inMbr  -ne "") { $mbr        = ($inMbr  -match '^[Yy]') }
                if ($inSort -ne "") { $sortByDate = ($inSort -match '^[Dd]') }
                Write-Host ""
                continue previewLoop
            }
            if ($pSel -eq 2) { Clear-Host; .\Main.ps1; return }
        } elseif ($pk.Key -eq [ConsoleKey]::Escape) {
            Clear-Host; .\Main.ps1; return
        }
    }
}

# -- Create Result folder and run ---------------------------------------------
[System.IO.Directory]::CreateDirectory($resultDir) | Out-Null

Write-Host ""
Write-Host "  Running DIA-NN ..." -ForegroundColor Cyan
Write-Host "  Output: $resultDir" -ForegroundColor DarkGray
Write-Host ""

& $diannCfg.DiannExe @diannArgs
$exitCode = $LASTEXITCODE

Write-Host ""
if ($exitCode -ne 0) {
    Write-Host "  DIA-NN exited with code $exitCode" -ForegroundColor Red
} else {
    Write-Host "  DIA-NN finished successfully." -ForegroundColor Green
    Write-Host "  Output: $resultDir" -ForegroundColor DarkGray
}

# -- Navigation ---------------------------------------------------------------
$nItems  = @("Open Result folder", "Re-run same folder", "Run analysis report (next step)", "Back to main menu")
$nColors = @("Cyan", "Cyan", "Cyan", "Cyan")
$nSel    = 0
Write-Host ""
Write-Host "  $rule" -ForegroundColor DarkCyan
$nTop = [Console]::CursorTop
for ($ni = 0; $ni -lt $nItems.Count; $ni++) {
    [Console]::SetCursorPosition(0, $nTop + $ni)
    if ($ni -eq 0) {
        Write-Host ("  > " + $nItems[$ni]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
    } else {
        Write-Host ("    " + $nItems[$ni]).PadRight($w + 4) -ForegroundColor $nColors[$ni] -NoNewline
    }
}
[Console]::SetCursorPosition(0, $nTop + $nItems.Count)
while ($true) {
    $k = [Console]::ReadKey($true)
    if ($k.Key -eq [ConsoleKey]::UpArrow -or $k.Key -eq [ConsoleKey]::DownArrow) {
        $p = $nSel
        if ($k.Key -eq [ConsoleKey]::UpArrow) { $nSel = ($nSel - 1 + $nItems.Count) % $nItems.Count }
        else { $nSel = ($nSel + 1) % $nItems.Count }
        [Console]::SetCursorPosition(0, $nTop + $p);    Write-Host ("    " + $nItems[$p]).PadRight($w + 4) -ForegroundColor $nColors[$p] -NoNewline
        [Console]::SetCursorPosition(0, $nTop + $nSel); Write-Host ("  > " + $nItems[$nSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
    } elseif ($k.Key -eq [ConsoleKey]::Enter) {
        if ($nSel -eq 0)      { Start-Process explorer.exe $resultDir }
        elseif ($nSel -eq 1)  { continue runLoop }
        elseif ($nSel -eq 2)  { Clear-Host; & ".\07_Report_generator.ps1"; return }
        else                  { Clear-Host; .\Main.ps1; return }
    } elseif ($k.Key -eq [ConsoleKey]::Escape) {
        Clear-Host; .\Main.ps1; return
    }
}

} # end :runLoop
