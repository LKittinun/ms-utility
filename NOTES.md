# Proteomics Process Suite - Development Notes

## CRITICAL: PowerShell String Rules (violations cause immediate crash)
- NEVER use em dash `--` in any string -> use plain hyphen `-`
- NEVER use Unicode box-drawing chars in strings -> use ASCII: `\--`, `+--`, `|   `
- Unicode in COMMENTS is fine; only strings/Write-Host lines are affected

---

## Script: 1_Project_init.ps1

### Flow Overview
1. Root directory (default Z:\Proteomics)
2. Analytics column ID (typed) -> $analyticsCol
3. Analytics column description (arrow-key library menu) -> $colDesc
4. Show existing projects under column
5. Project name -> $projectPath
6. Load $existingInfo from project_info.json if project already exists
7. PI name
8. Trap column ID (typed) -> $trapCol
9. Trap column description (arrow-key library menu, only if trapCol entered) -> $trapColDesc
10. Sample subfolders (with duplicate + conflict checks)
11. ProjectNo / ProjectID (preserved from JSON if existing project)
12. Preview + confirm
13. Create folders (CreateDirectory is idempotent)
14. Write project_info.json
15. Append or update column_log.csv

### Column ID vs Description (separated 2026-03-01)
- Column ID (e.g. C20533039) is always typed manually via Read-Host
- Column description (e.g. "Waters BEH C18 75um x 25cm") is selected from a
  library via arrow-key menu
- Analytics library: data\columns.json (string array)
- Trap library:      data\trap_columns.json (string array)
- Both libraries support:
  - Up/Down: navigate
  - Del: remove entry from library
  - Enter on [+ Add new description]: prompts, saves, re-shows menu
  - Enter on item: selects it
  - Blank library -> goes straight to add-new prompt
  - If add-new input is blank -> $colDesc = "" (no library change)
- DrawDescItem function defined once at script scope, reused for both menus
- Labeled loops: :descLoop (analytics), :trapDescLoop (trap)

### data\columns.json Format
Old format (migrated automatically on load):
  [{ "ColumnID": "C20533039", "Description": "Waters BEH C18 75um x 25cm" }]
New format (string array):
  ["Waters BEH C18 75um x 25cm", "Thermo PepMap C18 75um x 50cm"]
Migration: if items are PSCustomObjects, extract .Description and re-save.

### Existing Project Detection
When project_info.json already exists at $projectPath:
- $existingInfo is loaded from JSON
- Message: "Existing project found - leave fields blank to keep current values."
- Each prompt shows "(current: <value>)" hint in DarkGray
- Blank input -> keeps existing value
- Trap description menu pre-selects existing value if still in library
- ProjectID and ProjectNo are preserved from JSON (not regenerated)
- CSV: existing row is updated in-place (not a new row appended)
  - Match on Project name within the column's log file

### Subfolder Validation (two layers)
1. Duplicates within new input:
   - Uses hashtable with .ToLower() keys for case-insensitive comparison
   - Error: "Duplicate subfolders not allowed: <names>"
2. Conflicts with existing JSON SampleFolders:
   - Uses -icontains for case-insensitive match (Windows filesystem is case-insensitive)
   - Error: "Already exist in this project: <names>"
   - Only triggered when $existingInfo is set
On success: $subfolders = existing + new (merged, complete list stored in JSON/CSV)
CreateDirectory is idempotent so re-creating existing folders is harmless.

---

## Key Data Files

| File | Format | Description |
|---|---|---|
| data\columns.json | String array | Analytics column description library |
| data\trap_columns.json | String array | Trap column description library |
| Z:\Proteomics\<ColID>\column_log.csv | CSV | Per-column project log |
| Z:\Proteomics\<ColID>\<Project>\project_info.json | JSON | Per-project metadata |

### project_info.json Fields
ProjectID, Project, PI, AnalyticsColumn, ColumnDescription,
TrapColumn, TrapColumnDescription, ProjectNo, Created (yyyy-MM-dd HH:mm),
SampleFolders (array)

### column_log.csv Fields
ProjectID, ProjectNo, Date, Project, PI, AnalyticsColumn, ColumnDescription,
TrapColumn, TrapColumnDescription, SampleFolders (semicolon-joined)

---

## Known Patterns / Gotchas

- ConvertFrom-Json on single-item array returns object not array
  -> always wrap with @()
- ConvertTo-Json on single-item array returns scalar
  -> always use: ConvertTo-Json -InputObject @($array)
- Get-ChildItem -Include without trailing \* silently skips root-level files
  -> use -Path "$path\*"
- ProjectNo counter only counts folders containing project_info.json
- [System.IO.Directory]::CreateDirectory() is idempotent (safe on existing folders)
- $existingInfo.PI can be $null (JSON null) if PI was not set
  -> if ($existingInfo.PI) { ... } handles both null and empty string

---

## Description Libraries (current entries)

data\columns.json (analytics):
  - "Thermo PepMap C18 75um x 50cm"
  - "Waters BEH C18 75um x 25cm"
  - "Ionopticks C18 75um x 25cm"

data\trap_columns.json (trap):
  - Created on first use; entries added interactively

---

## Verified

| Script | Status | Notes |
|---|---|---|
| 1_Project_init.ps1 | VERIFIED 2026-03-01 | Full new flow with all changes |
| 2_Repair_project_order.ps1 | VERIFIED | Re-numbers C20531700 (8 projects) |
| 3_Backfill_column.ps1 | VERIFIED | Retroactive metadata for C20531700 |
| Main.ps1 | VERIFIED | Grouped menu, non-selectable separators |
