Write-Host "
** Process basic metrics from DIA-NN output **
"

$first_path = Get-location
$path = Read-Host "Insert directory, leave blank for a current location"

$path = $path.Replace("\","/")

& "C:\Program Files\R\R-4.4.1\bin\Rscript.exe" ".\R\Data.R" $path 2>&1| %{"$_"}

$Which = Read-Host '
Another process?
y = Menu, other = Exit'

switch ($Which) {
  y {.\Main.ps1}
  default {'Exiting...'}
}