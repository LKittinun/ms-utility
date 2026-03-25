# Mass Spectrometry Utility Suite

A collection of PowerShell scripts for managing LC-MS/MS proteomics projects, column logs, DIA-NN metrics, and service reports.

## Requirements

- Windows PowerShell 5.1+
- Network access to the proteomics data root (default: `Z:\Proteomics`)
- [ImportExcel](https://github.com/dfinke/ImportExcel) module (auto-installed on first use by scripts that need it)
- R + required packages (for script 06, report generation)
- [msConvert](https://proteowizard.sourceforge.io/) on PATH (for script 07)
- [mzsniffer](https://github.com/LewisResearchGroup/mzsniffer) (for script 08)

## Usage

Launch the interactive menu:

```powershell
.\Main.ps1
```

Navigate with Up/Down arrows and Enter. The root directory is shown in the footer and can be changed via **Settings > Set root directory**.

## Scripts

| # | Script | Section | Description |
|---|--------|---------|-------------|
| 01 | `01_Project_init.ps1` | Project | Create or update a project: folders, metadata JSON, column log CSV |
| 02 | `02_Find_project.ps1` | Project | Multi-term search by column, name, PI, or project ID; open in Explorer |
| 03 | `03_Projects_overview.ps1` | Project | Filter all projects and export to Excel or CSV |
| 04 | `04_Column_usage.ps1` | Analysis | Column usage report across projects |
| 05 | `05_DIANN_metrics.ps1` | Analysis | DIA-NN metrics plots and TSV export |
| 06 | `06_Report_generator.ps1` | Analysis | Generate service report Excel workbook per sample |
| 07 | `07_Bulk_msConvert.ps1` | Miscellaneous | Bulk convert `.raw` files to mzML |
| 08 | `08_Contaminant_check.ps1` | Miscellaneous | Contaminant screening via mzsniffer |
| 09 | `09_Clear_files.ps1` | Miscellaneous | Delete `.sld` and `.meth` files |
| 10 | `10_Repair_project_order.ps1` | Admin | Re-number projects by creation date |
| 11 | `11_Backfill_column.ps1` | Admin | Backfill column log from existing project folders |
| 12 | `12_Sync_from_overview.ps1` | Admin | Apply edits from overview CSV back to project JSON files |

Admin scripts (10-12) are password-protected.

## Directory Structure

```
Z:\Proteomics\Projects\
  YYYY-MM-DD_<ColumnID>\
    column_log.csv
    YYYY-MM-DD_<ProjectName>\
      project_info.json
      <SampleFolder>\
        Result\          <- DIA-NN output
```

## Configuration

`config.json` in the script root sets the data root:

```json
{ "Root": "Z:\\Proteomics" }
```

Change it from the menu (Settings > Set root directory) or edit the file directly.
