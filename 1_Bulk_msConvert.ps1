Write-Host "
** Convert bulk .raw file to mzML **
"

$first_path = Get-Location
$path = Read-Host "Set the path to the directory containing the .raw files, leave blank for a current location"
Set-location -Path $path

$demux = Read-Host "Demultiplex? y = yes, otherwise = no"
# Set the output directory (optional, if you want to specify a different output directory)
$outputDir = [System.IO.Path]::Combine($path, "mzML_files") 
Write-Host $outputDir

# Create the output directory if it does not exist
if (-Not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir
}

# Get all .raw files in the directory and convert if not already converted
Get-ChildItem -Filter *.raw | ForEach-Object {
    $outputFile = Join-Path $outputDir "$($_.BaseName).mzML"
    if (Test-Path $outputFile) {
        Write-Host "File $outputFile already exists. Skipping conversion."
    }
    elseif ($demux -eq "y") {
        msconvert $_.FullName --mzML --outdir $outputDir  --zlib --filter "peakPicking vendor msLevel=1-" --filter "zeroSamples removeExtra 1-" --filter "demultiplex optimization=overlap_only" -v
    } 
    else {
        msconvert $_.FullName --mzML --outdir $outputDir  --zlib --filter "peakPicking vendor msLevel=1-" --filter "zeroSamples removeExtra 1-" -v
    } 
}

Write-Host "Conversion process complete."

Set-Location -Path $first_path

$Which = Read-Host '
Another process?
y = Menu, other = Exit'

switch ($Which) {
  y {.\Main.ps1}
  default {'Exiting...'}
}