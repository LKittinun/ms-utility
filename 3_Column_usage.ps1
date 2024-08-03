Write-Host "
** Check the total column usage. All files must be within a column's parent directory **
"

$first_path = Get-location
$path = Read-Host "Insert directory, leave blank for a current location"
Set-location -Path $path

# Set the output directory (optional, if you want to specify a different output directory)
$currentDate = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFilePath = [System.IO.Path]::Combine($path, "Column_usage_history\")
$outputCsvFilePath = "$outputFilePath\Column_info_$currentdate.csv"
$outputTextFilePath = "$outputFilePath\Column_info_$currentdate.txt"

if (-Not (Test-Path -Path $outputFilePath)) {
    New-Item -Path $outputFilePath -ItemType Directory > $null
}

# Get all files recursively and sort by creation date
$files = Get-ChildItem *.raw -Path $path -File -Recurse | Sort-Object CreationTime

$fileInfoList = @()

# Loop through each file and collect information
foreach ($file in $files) {
    $fileInfoList += [PSCustomObject]@{
        Name        = $file.Name
        FullPath    = $file.FullName
        Size        = $file.Length
        CreationTime = $file.CreationTime
    }
}

# Display the file information list

$fileInfoList | Export-Csv -Path $outputCsvFilePath -NoTypeInformation

# Calculate additional information
$totalFiles = $fileInfoList.Count
$firstFile = $files | Select-Object -First 1
$lastFile = $files | Select-Object -Last 1

# Create or clear the output text file
New-Item -Path $outputTextFilePath -ItemType File -Force > $null

# Write summary information to the text file
Add-Content -Path $outputTextFilePath -Value "Total runs: $totalFiles"
Add-Content -Path $outputTextFilePath -Value "First run: $($firstFile.CreationTime) - $($firstFile.FullName)"
Add-Content -Path $outputTextFilePath -Value "Last run: $($lastFile.CreationTime) - $($lastFile.FullName)"

Write-Host "
--------------------------------"
Get-content $outputTextFilePath | Write-Host
Write-Host "--------------------------------"

Write-Host "
Data is created in $outputFilePath"

Set-location $first_path

$Which = Read-Host '
Another process?
y = Menu, other = Exit'

switch ($Which) {
  y {.\Main.ps1}
  default {'Exiting...'}
}