$path = Read-Host "Insert directory, leave blank for a current location"
$files = Get-ChildItem -Path $path -Recurse -Include *.sld, *.meth -File


if ($files.Count -eq 0) {
    Write-Host "
No *sld and *meth files found"
   }
else {
   Write-Host "
These files will be removed:"
   $files

   $confirm = Read-Host "
Confirm? = y"
   switch ($confirm){
   y {rm $files
      Write-Host "
All files are removed."}
   default {Write-Host "
Cancelled"}
    }
}

$Which = Read-Host '
Another process?
y = Menu, other = Exit'

switch ($Which) {
  y {.\Main.ps1}
  default {'Exiting...'}
}