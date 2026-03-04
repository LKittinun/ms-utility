$w      = 55
$border = "=" * $w
$rule   = "-" * $w
$root   = "Z:\Proteomics\Projects"

function Show-Header {
    Clear-Host
    Write-Host ""
    Write-Host "  $border" -ForegroundColor DarkCyan
    Write-Host "   [4]  Find project" -ForegroundColor Cyan
    Write-Host "        Search by ID, project name, or PI" -ForegroundColor DarkCyan
    Write-Host "  $border" -ForegroundColor DarkCyan
    Write-Host ""
}

Show-Header

while ($true) {
    $raw = (Read-Host "  Search (blank = main menu)").Trim()
    if ($raw -eq "") { Clear-Host; .\Main.ps1; return }

    $terms = @($raw -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })

    Write-Host ""
    Write-Host "  Searching $root ..." -ForegroundColor DarkCyan

    $jsonFiles = @(Get-ChildItem -Path $root -Recurse -Filter "project_info.json" -ErrorAction SilentlyContinue)

    $results = @()
    foreach ($jf in $jsonFiles) {
        try { $info = Get-Content $jf.FullName -Raw | ConvertFrom-Json } catch { continue }

        $match = $false
        foreach ($term in $terms) {
            if (($info.ProjectID -and $info.ProjectID -like "*$term*") -or
                ($info.Project   -and $info.Project   -like "*$term*") -or
                ($info.PI        -and $info.PI        -like "*$term*")) {
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

    $nSel = 0
    $nTop = [Console]::CursorTop

    for ($i = 0; $i -lt $nItems.Count; $i++) {
        [Console]::SetCursorPosition(0, $nTop + $i)
        $fg = if ($nItems[$i] -eq "Exit") { "DarkYellow" } else { "DarkCyan" }
        Write-Host ("    " + $nItems[$i]).PadRight($w + 4) -ForegroundColor $fg -NoNewline
    }
    [Console]::SetCursorPosition(0, $nTop)
    Write-Host ("  > " + $nItems[0]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
    [Console]::SetCursorPosition(0, $nTop + $nItems.Count)

    $action = $null
    while ($null -eq $action) {
        $k = [Console]::ReadKey($true)
        if ($k.Key -eq [ConsoleKey]::UpArrow -or $k.Key -eq [ConsoleKey]::DownArrow) {
            $p    = $nSel
            $nSel = if ($k.Key -eq [ConsoleKey]::UpArrow) {
                        ($nSel - 1 + $nItems.Count) % $nItems.Count
                    } else {
                        ($nSel + 1) % $nItems.Count
                    }
            [Console]::SetCursorPosition(0, $nTop + $p)
            $fg = if ($nItems[$p] -eq "Exit") { "DarkYellow" } else { "DarkCyan" }
            Write-Host ("    " + $nItems[$p]).PadRight($w + 4) -ForegroundColor $fg -NoNewline
            [Console]::SetCursorPosition(0, $nTop + $nSel)
            Write-Host ("  > " + $nItems[$nSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
            [Console]::SetCursorPosition(0, $nTop + $nItems.Count)
        } elseif ($k.Key -eq [ConsoleKey]::Enter) {
            if ($nItems[$nSel] -eq "Open in Explorer") {
                if ($results.Count -eq 1) {
                    explorer.exe $results[0].Path
                } else {
                    # Secondary arrow-key picker
                    $pItems = @($results | ForEach-Object {
                        ("$($_.ProjectID)  $($_.Project)").PadRight(40).Substring(0, 40)
                    })
                    $pSel = 0
                    Write-Host ""
                    Write-Host "  Open which project?" -ForegroundColor DarkCyan
                    $pTop = [Console]::CursorTop

                    for ($i = 0; $i -lt $pItems.Count; $i++) {
                        [Console]::SetCursorPosition(0, $pTop + $i)
                        Write-Host ("    " + $pItems[$i]).PadRight($w + 4) -ForegroundColor DarkCyan -NoNewline
                    }
                    [Console]::SetCursorPosition(0, $pTop)
                    Write-Host ("  > " + $pItems[0]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
                    [Console]::SetCursorPosition(0, $pTop + $pItems.Count)

                    $picked = $null
                    while ($null -eq $picked) {
                        $pk = [Console]::ReadKey($true)
                        if ($pk.Key -eq [ConsoleKey]::UpArrow -or $pk.Key -eq [ConsoleKey]::DownArrow) {
                            $pp   = $pSel
                            $pSel = if ($pk.Key -eq [ConsoleKey]::UpArrow) {
                                        ($pSel - 1 + $pItems.Count) % $pItems.Count
                                    } else {
                                        ($pSel + 1) % $pItems.Count
                                    }
                            [Console]::SetCursorPosition(0, $pTop + $pp)
                            Write-Host ("    " + $pItems[$pp]).PadRight($w + 4) -ForegroundColor DarkCyan -NoNewline
                            [Console]::SetCursorPosition(0, $pTop + $pSel)
                            Write-Host ("  > " + $pItems[$pSel]).PadRight($w + 4) -ForegroundColor Black -BackgroundColor Cyan -NoNewline
                            [Console]::SetCursorPosition(0, $pTop + $pItems.Count)
                        } elseif ($pk.Key -eq [ConsoleKey]::Enter) {
                            $picked = $results[$pSel].Path
                        } elseif ($pk.Key -eq [ConsoleKey]::Escape) {
                            $picked = ""
                        }
                    }
                    if ($picked -ne "") { explorer.exe $picked }
                }
                # stay in nav loop to allow another action
            } else {
                $action = $nItems[$nSel]
            }
        } elseif ($k.Key -eq [ConsoleKey]::Escape) {
            $action = "Exit"
        }
    }

    if ($action -eq "Back to main menu") { Clear-Host; .\Main.ps1; return }
    if ($action -eq "Exit") {
        [Console]::SetCursorPosition(0, $nTop + $nItems.Count + 1)
        Write-Host "  Exiting..." -ForegroundColor DarkYellow
        return
    }
    # "New search"
    Show-Header
}
