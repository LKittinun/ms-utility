$w          = 55
$border     = "=" * $w
$rule       = "-" * $w
. (Join-Path $PSScriptRoot "lib\Menu.ps1")
$_cfg       = if (Test-Path (Join-Path $PSScriptRoot "config.json")) { Get-Content (Join-Path $PSScriptRoot "config.json") -Raw | ConvertFrom-Json } else { $null }
$_rootBase  = if ($_cfg -and $_cfg.Root) { $_cfg.Root } else { "Z:\Proteomics" }
$root       = Join-Path $_rootBase "Projects"
$prohibited = @("blank", "raw_summary", "prtc", "sst", "column_usage_history")

function Show-Header {
    Clear-Host
    Write-Host ""
    Write-Host "  $border" -ForegroundColor DarkCyan
    Write-Host "   [2]  Find project" -ForegroundColor Cyan
    Write-Host "        Search by column, project name, PI, or project ID" -ForegroundColor DarkCyan
    Write-Host "  $border" -ForegroundColor DarkCyan
    Write-Host ""
}

Show-Header

while ($true) {
    $raw = (Read-Host "  Search (blank = main menu)").Trim()
    if ($raw -eq "") { Return-ToMain; return }

    $terms = @($raw -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })

    Write-Host ""
    Write-Host "  Searching $root ..." -ForegroundColor DarkCyan

    if (-not (Test-Path $root)) {
        Write-Host "  Path not found: $root" -ForegroundColor Red
        Write-Host "  Use 'Set root directory' in the main menu to fix this." -ForegroundColor DarkGray
        Write-Host ""
        continue
    }

    $jsonFiles = @(Get-ChildItem -Path $root -Recurse -Filter "project_info.json" -ErrorAction SilentlyContinue)
    Write-Host "  Found $($jsonFiles.Count) project_info.json file(s)" -ForegroundColor DarkGray
    $jsonFiles = @($jsonFiles | Where-Object {
        $prohibited -notcontains ($_.Directory.Name -replace '^\d{4}-\d{2}-\d{2}_','').ToLower()
    })

    $results = @()
    foreach ($jf in $jsonFiles) {
        try { $info = Get-Content $jf.FullName -Raw | ConvertFrom-Json } catch { continue }

        $match = $false
        foreach ($term in $terms) {
            if (($info.ProjectID       -and $info.ProjectID       -like "*$term*") -or
                ($info.Project         -and $info.Project         -like "*$term*") -or
                ($info.PI              -and $info.PI              -like "*$term*") -or
                ($info.AnalyticsColumn -and $info.AnalyticsColumn -like "*$term*") -or
                ($info.TrapColumn      -and $info.TrapColumn      -like "*$term*")) {
                $match = $true; break
            }
        }

        if ($match) {
            $results += [PSCustomObject]@{
                ProjectID = if ($info.ProjectID)       { $info.ProjectID }       else { "-" }
                Project   = if ($info.Project)         { $info.Project   }       else { "-" }
                PI        = if ($info.PI)              { $info.PI        }       else { "-" }
                Column    = if ($info.AnalyticsColumn) { $info.AnalyticsColumn } else { "-" }
                Path      = $jf.DirectoryName
            }
        }
    }

    $termList = ($terms | ForEach-Object { "'$_'" }) -join ", "
    Write-Host ""
    if ($results.Count -eq 0) {
        Write-Host "  No results for $termList" -ForegroundColor Yellow
    } else {
        Write-Host "  $($results.Count) result(s) for $termList" -ForegroundColor Cyan
        Write-Host ""
        foreach ($r in $results) {
            Write-Host "  $rule" -ForegroundColor DarkCyan
            Write-Host "  ID      : " -NoNewline -ForegroundColor DarkCyan
            Write-Host $r.ProjectID -ForegroundColor White
            Write-Host "  Project : " -NoNewline -ForegroundColor DarkCyan
            Write-Host $r.Project -ForegroundColor White
            Write-Host "  PI      : " -NoNewline -ForegroundColor DarkCyan
            Write-Host $r.PI -ForegroundColor White
            Write-Host "  Column  : " -NoNewline -ForegroundColor DarkCyan
            Write-Host $r.Column -ForegroundColor White
            Write-Host "  Path    : " -NoNewline -ForegroundColor DarkCyan
            Write-Host $r.Path -ForegroundColor Cyan
        }
        Write-Host "  $rule" -ForegroundColor DarkCyan
    }
    Write-Host ""

    # -- Navigation ------------------------------------------------------------
    $nItems = @()
    if ($results.Count -ge 1) { $nItems += "Open in Explorer" }
    $nItems += "New search"
    $nItems += "Back to main menu"
    $nItems += "Exit"

    $action = $null
    while ($null -eq $action) {
        $nr = Show-Menu -Items $nItems -AllowEscape
        if ($nr.Action -ne "select") {
            $action = "Exit"
        } elseif ($nItems[$nr.Index] -eq "Open in Explorer") {
            if ($results.Count -eq 1) {
                explorer.exe $results[0].Path
            } else {
                # Secondary picker
                $pItems = @($results | ForEach-Object {
                    ("$($_.ProjectID)  $($_.Project)").PadRight(40).Substring(0, 40)
                })
                Write-Host ""
                Write-Host "  Open which project?" -ForegroundColor DarkCyan
                $pr = Show-Menu -Items $pItems -AllowEscape
                if ($pr.Action -eq "select") { explorer.exe $results[$pr.Index].Path }
            }
            # stay in nav loop to allow another action
        } else {
            $action = $nItems[$nr.Index]
        }
    }

    if ($action -eq "Back to main menu") { Return-ToMain; return }
    if ($action -eq "Exit") { Request-Exit; return }
    # "New search"
    Show-Header
}
