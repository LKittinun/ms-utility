Write-Host "
** Mass spectrometry utility script **
"

$Which = Read-Host 'What would you like to do?

1: Bulk convert to mzML (msConvert)
2: Check contaminant (mzsniffer)
3: Column usage
4: DIA-NN basic data process
99: Clean up temporary method files (*sld, *meth)

Other: Exit

Your choice'

switch ($Which) {
  1 {.\1_Bulk_msConvert.ps1}
  2 {.\2_Contaminant_check.ps1}
  3 {.\3_Column_usage.ps1}
  4 {.\4_DIANN.ps1}
  99 {.\99_Clear_files.ps1}
  default {'Exiting...'}
}