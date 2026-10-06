# Shared name-matching helpers (PI names etc.).
# Dot-source from a script:  . (Join-Path $PSScriptRoot "lib\Names.ps1")

# Collapse whitespace + lowercase, for comparing names
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

# Existing name that a typed name probably means, or $null:
#  1. an existing name appears as a whole word in the typed name
#     (e.g. "Assoc. Prof. Dr. Raphatphorn Navakanitworakul" -> "Raphatphorn")
#  2. closest existing name within 2 edits (1 edit for names under 5 letters)
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
