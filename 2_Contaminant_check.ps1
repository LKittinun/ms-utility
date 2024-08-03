write-host "
** Check contaminant with mzsniffer. For mzML files only** 
" 
$first_path = Get-location 
$path = Read-Host "Insert directory, leave blank for a current location"
Set-location -Path $path
$files = Get-ChildItem "*.mzML" | Sort-Object LastWriteTime
$logFilePath = ".\contaminant_check.txt"
$mzsnifferPath = Join-Path -Path $first_path -ChildPath "mzsniffer\mzsniffer.exe"

Write-Host "Start a checking process ..."
    
if (Test-Path $logFilePath) {
     Write-Warning "An old file existed. Running this script will replace an existing file."
     Remove-Item $logFilePath -Force
     }

foreach ($file in $files) {
    $d = [datetime](Get-ItemProperty -Path $file -Name LastWriteTime).lastwritetime
    write-host "Processing file $file, last written: $d." -InformationVariable message
    $output = & $mzsnifferPath $file.FullName 2>&1
    $filteredOutput = $output -split "\r?\n" | Where-Object { $_ -notmatch "^At|^    \+ CategoryInfo|^    \+ FullyQualifiedErrorId|^    \+ \d" }

    Write-Output $filteredOutput
    $message | Out-File -FilePath $logFilePath -Append
    $filteredOutput | Add-Content -Path $logFilePath
    Add-Content -Path $logFilePath -Value "`n------- `n"
}

Write-Host "mzsniffer process finished. Check the log file '$logFilePath' for details."
Set-location $first_path

$Which = Read-Host '
Another process? 
y = Menu, other = Exit'

switch ($Which) {
  y {.\Main.ps1}
  default {'Exiting...'}
}