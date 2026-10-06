#!/usr/bin/env Rscript
# =============================================================================
# Proteomics Core Facility Report Generator  - DIA-NN Output
# =============================================================================
# Generates one Excel file from a DIA-NN project folder:
#
#   Analysis_Report.xlsx  - overview + methods, QC sheets, figures, and the
#                           pg_matrix (with contaminant flags)
#
# Expected folder structure
# -------------------------
#   <project_folder>/                  <- pass THIS path as the argument
#   +-- *.raw                          Thermo raw MS files
#   +-- *.raw.quant                    DIA-NN quantification cache
#   +-- Result/                        DIA-NN output subfolder
#       +-- report.log.txt
#       +-- report.stats.tsv
#       +-- report.<name>.pg_matrix.tsv   primary result
#       +-- report.pr_matrix.tsv
#       +-- report.gg_matrix.tsv
#       +-- ...
#
# Usage (called automatically by 5_Report_generator.ps1)
# -------------------------------------------------------
#   Rscript generate_report.R  <project_dir>
#   Rscript generate_report.R  <project_dir>  <result_subdir>
#   Rscript generate_report.R  <project_dir>  <result_subdir>  <output_dir>
# =============================================================================

# --- Package bootstrap --------------------------------------------------------
for (pkg in c("openxlsx", "tools")) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    message(sprintf("Installing missing package: %s", pkg))
    install.packages(pkg, repos = "https://cloud.r-project.org", quiet = TRUE)
  }
  suppressPackageStartupMessages(library(pkg, character.only = TRUE))
}

# --- Configuration ------------------------------------------------------------
RESULT_DIR_NAME <- "Result"

# Signature shown on the report ("Prepared by") and stored as the workbook author
REPORT_AUTHOR      <- "Kittinun Leetanaporn"
REPORT_AFFILIATION <- "Proteomics Core Facility"
RAW_EXTENSIONS  <- c(".raw")

# --- Theme (workbook + figures) -----------------------------------------------
CLR_DARK_BLUE  <- "#1F3864"   # navy - titles, table headers
CLR_MID_BLUE   <- "#2F5496"   # secondary headers, section text
CLR_LIGHT_BLUE <- "#F2F4F8"   # table row banding
CLR_WHITE      <- "#FFFFFF"
CLR_GOOD       <- "#E2EFDA"
CLR_WARN       <- "#FFF2CC"
CLR_BAD        <- "#FCE4E4"
CLR_SECTION    <- "#E9EDF4"   # section header band
CLR_RULE       <- "#BFC9D9"   # thin borders / rules
CLR_TEXT_MUTED <- "#595959"
BASE_FONT      <- "Calibri"

# Figure palette
FIG_MAIN   <- "#2F5496"
FIG_DARK   <- "#1F3864"
FIG_ACCENT <- "#C55A11"
FIG_MUTED  <- "#A6A6A6"
FIG_GRID   <- "#E7E7E7"
FIG_TEXT   <- "#404040"
# Cairo gives clean greyscale anti-aliasing (the Windows device adds colour fringes)
PNG_TYPE   <- if (isTRUE(capabilities("cairo"))) "cairo" else getOption("bitmapType", "windows")

PG_META <- c("Protein.Group", "Protein.Ids", "Protein.Names", "Genes",
             "First.Protein.Description", "N.Sequences",
             "N.Proteotypic.Sequences")

# Fallback patterns for contaminant protein groups (in addition to the
# --cont-quant-exclude tag found in the DIA-NN log)
CONTAM_PATTERNS <- c("cRAP", "^CON__", "Cont_")

# --- Arguments ----------------------------------------------------------------
args <- commandArgs(trailingOnly = TRUE)

project_dir_arg <- if (length(args) >= 1) args[1] else getwd()
result_dir_name <- if (length(args) >= 2) args[2] else RESULT_DIR_NAME
output_dir_arg  <- if (length(args) >= 3) args[3] else project_dir_arg

project_dir <- normalizePath(project_dir_arg, winslash = "/", mustWork = FALSE)
result_dir  <- file.path(project_dir, result_dir_name)
output_dir  <- normalizePath(output_dir_arg, winslash = "/", mustWork = FALSE)
second_dir_arg <- if (length(args) >= 4) args[4] else NULL
second_dir <- if (!is.null(second_dir_arg) && nchar(second_dir_arg) > 0) {
  normalizePath(second_dir_arg, winslash = "/", mustWork = FALSE)
} else NULL

sep <- strrep("=", 62)
cat(sprintf("\n%s\n", sep))
cat(" Proteomics Report Generator  |  DIA-NN Output\n")
cat(sprintf("%s\n", sep))
cat(sprintf(" Project : %s\n", project_dir))
cat(sprintf(" Results : %s\n", result_dir))
cat(sprintf(" Output  : %s\n", output_dir))
if (!is.null(second_dir)) cat(sprintf(" Extra   : %s\n", second_dir))
cat(sprintf("%s\n\n", sep))

# -----------------------------------------------------------------------------
# UTILITY FUNCTIONS
# -----------------------------------------------------------------------------

bytes_to_human <- function(n) {
  units <- c("B", "KB", "MB", "GB", "TB")
  for (u in units) {
    if (abs(n) < 1024) return(sprintf("%.2f %s", n, u))
    n <- n / 1024
  }
  sprintf("%.2f PB", n)
}

# Sample columns = numeric columns that are not annotation columns
get_sample_cols <- function(df) {
  cand <- setdiff(names(df), c(PG_META, "Contaminant"))
  cand[vapply(df[cand], is.numeric, logical(1))]
}

# "D:/runs/QC_01.raw" -> "QC_01"
shorten_name <- function(n) {
  if (grepl("[/\\\\]", n)) tools::file_path_sans_ext(basename(n)) else n
}

# TRUE for contaminant protein groups (log tag or common contaminant prefixes)
flag_contaminants <- function(pg_df, cont_tag = "n/a") {
  pats <- CONTAM_PATTERNS
  if (!is.null(cont_tag) && !identical(cont_tag, "n/a") && nchar(cont_tag) > 0)
    pats <- c(gsub("([][{}()+*^$|\\\\?.])", "\\\\\\1", cont_tag), pats)
  rx <- paste(pats, collapse = "|")
  ids <- if ("Protein.Group" %in% names(pg_df)) pg_df$Protein.Group else rep("", nrow(pg_df))
  grepl(rx, ids, ignore.case = TRUE, perl = TRUE)
}

# Wrapped, merged paragraph across cols 1:2; returns next free row
write_paragraph <- function(wb, sheet, text, row, font_size = 10, italic = FALSE,
                            colour = "#000000", chars_per_line = 120) {
  writeData(wb, sheet, text, startRow = row, startCol = 1)
  addStyle(wb, sheet,
           createStyle(wrapText = TRUE, fontSize = font_size, valign = "top",
                       fontColour = colour,
                       textDecoration = if (italic) "italic" else NULL),
           rows = row, cols = 1:2, gridExpand = TRUE, stack = TRUE)
  mergeCells(wb, sheet, cols = 1:2, rows = row)
  n_lines <- max(1, ceiling(nchar(text) / chars_per_line))
  setRowHeights(wb, sheet, rows = row, heights = max(15, n_lines * (font_size + 4)))
  row + 1
}

# Section header row in the house style; returns next free row
write_section <- function(wb, sheet, title, row) {
  writeData(wb, sheet, title, startRow = row, startCol = 1)
  addStyle(wb, sheet,
           createStyle(fontColour = CLR_MID_BLUE, textDecoration = "bold",
                       fgFill = CLR_SECTION, fontSize = 12),
           rows = row, cols = 1:2, gridExpand = TRUE, stack = TRUE)
  mergeCells(wb, sheet, cols = 1:2, rows = row)
  row + 1
}

# -----------------------------------------------------------------------------
# DATA COLLECTION
# -----------------------------------------------------------------------------

collect_raw_files <- function(project_dir, second_dir = NULL) {
  find_raw <- function(dir) {
    files <- character(0)
    for (ext in RAW_EXTENSIONS)
      files <- c(files, list.files(dir, pattern = paste0("\\", ext, "$"),
                                   full.names = TRUE, ignore.case = TRUE))
    sort(files)
  }

  # second_dir is already resolved by the caller (07_Report_generator.ps1) to
  # the exact folder to search; it is used as-is
  two_sources <- !is.null(second_dir) && dir.exists(second_dir) &&
    normalizePath(second_dir, winslash = "/") != normalizePath(project_dir, winslash = "/")

  # Label each source by folder name; use the full path when the name is empty
  # (drive root) or both sources share the same name
  lbl1 <- basename(project_dir)
  lbl2 <- if (two_sources) basename(second_dir) else ""
  if (two_sources && (lbl2 == "" || tolower(lbl2) == tolower(lbl1))) lbl2 <- second_dir
  if (lbl1 == "") lbl1 <- project_dir

  dirs <- if (two_sources) {
    list(list(path = project_dir, label = lbl1),
         list(path = second_dir,  label = lbl2))
  } else {
    list(list(path = project_dir, label = NULL))
  }

  seen <- character(0)
  all_rows <- lapply(dirs, function(d) {
    files <- find_raw(d$path)
    # Skip files already found in an earlier source (e.g. partially archived copy)
    dup   <- tolower(basename(files)) %in% seen
    if (any(dup))
      cat(sprintf("  [INFO] %d raw file(s) in %s already listed from project folder - skipped\n",
                  sum(dup), d$path))
    files <- files[!dup]
    seen <<- c(seen, tolower(basename(files)))
    if (length(files) == 0) return(NULL)
    rows <- lapply(files, function(f) {
      st  <- file.info(f)
      row <- data.frame(`File Name` = basename(f), check.names = FALSE, stringsAsFactors = FALSE)
      if (two_sources) row[["Source Folder"]] <- d$label
      row[["Size (MB)"]] <- round(st$size / 1024^2, 2)
      row[["Size (GB)"]] <- round(st$size / 1024^3, 3)
      # mtime: on Windows ctime is when the file was copied/archived, while the
      # last-modified time stays close to the end of acquisition
      row[["File Date"]] <- format(st$mtime, "%Y-%m-%d %H:%M")
      row
    })
    do.call(rbind, rows)
  })
  all_rows <- Filter(Negate(is.null), all_rows)
  if (length(all_rows) == 0) return(data.frame())
  do.call(rbind, all_rows)
}

parse_log <- function(log_path) {
  info <- list(
    # --- System ----------------------------------------------------------
    "DIA-NN Version"                = "n/a",
    # --- Analysis settings -----------------------------------------------
    "FDR Threshold (q-value)"       = "n/a",
    "Quantification Method"         = "n/a",
    "Match-Between-Runs (MBR)"      = "n/a",
    "Output Matrices"               = "n/a",
    "Generate Spectral Library"     = "n/a",
    # --- Library / database ----------------------------------------------
    "Spectral Library"              = "n/a",
    "FASTA Database(s)"             = "n/a",
    "Contaminant Exclusion Tag"     = "n/a",
    # --- Peptide settings ------------------------------------------------
    "Peptide Length"                = "n/a",
    "Precursor m/z"                 = "n/a",
    "Precursor Charge"              = "n/a",
    "Fragment m/z"                  = "n/a",
    "Enzyme Cut Sites"              = "n/a",
    "Missed Cleavages"              = "n/a",
    "N-term Met Excision"           = "n/a",
    "Fixed Modification (Cys)"      = "n/a",
    "Variable Modifications"        = "n/a",
    "Max Variable Mods / Peptide"   = "n/a",
    "Cross-run Normalisation"       = "n/a",
    # --- Mass accuracy ---------------------------------------------------
    "MS2 Mass Accuracy (ppm)"       = "n/a",
    "MS1 Mass Accuracy (ppm)"       = "n/a",
    # --- Results ---------------------------------------------------------
    "Input Raw Files (count)"       = "n/a",
    "Protein Groups (q <= 0.01)"    = "n/a"
  )
  if (!file.exists(log_path)) return(info)

  text <- tryCatch(
    paste(readLines(log_path, warn = FALSE, encoding = "UTF-8"), collapse = "\n"),
    error = function(e) ""
  )
  if (nchar(text) == 0) return(info)

  rx <- function(pattern, default = "n/a") {
    m <- regmatches(text, regexpr(pattern, text, perl = TRUE))
    if (length(m) == 0 || nchar(m) == 0) return(default)
    trimws(sub(pattern, "\\1", m, perl = TRUE))
  }

  ver_str <- rx("DIA-NN\\s+([\\d.]+[^\\r\\n]*?)\\n")
  info[["DIA-NN Version"]]             <- ver_str

  info[["FDR Threshold (q-value)"]]    <- rx("--qvalue\\s+([\\d.]+)", default = "0.01 (DIA-NN default)")
  major_ver <- suppressWarnings(as.integer(sub("^(\\d+)\\..*", "\\1", ver_str)))
  info[["Quantification Method"]]      <- if (grepl("--direct-quant", text, fixed = TRUE)) "Legacy"
                                          else if (!is.na(major_ver) && major_ver >= 2) "QuantUMS"
                                          else "MaxLFQ"
  info[["Match-Between-Runs (MBR)"]]   <- ifelse(grepl("--reanalyse",    text), "Yes", "No")
  info[["Output Matrices"]]            <- ifelse(grepl("--matrices",     text), "Yes", "No")
  info[["Generate Spectral Library"]]  <- ifelse(grepl("--gen-spec-lib", text), "Yes", "No")

  m_lib <- regmatches(text, regexpr("--lib\\s+(\\S+)", text, perl = TRUE))
  if (length(m_lib) > 0)
    info[["Spectral Library"]] <- basename(sub("--lib\\s+(\\S+)", "\\1", m_lib, perl = TRUE))

  fastas <- unlist(regmatches(text, gregexpr("(?<=--fasta\\s)\\S+", text, perl = TRUE)))
  if (length(fastas) > 0)
    info[["FASTA Database(s)"]] <- paste(basename(fastas), collapse = "; ")

  info[["Contaminant Exclusion Tag"]]  <- rx("--cont-quant-exclude\\s+(\\S+)")

  pep_min <- rx("--min-pep-len\\s+(\\d+)"); pep_max <- rx("--max-pep-len\\s+(\\d+)")
  if (pep_min != "n/a" || pep_max != "n/a") info[["Peptide Length"]] <- paste0(pep_min, "-", pep_max, " aa")

  pr_mz_min <- rx("--min-pr-mz\\s+([\\d.]+)"); pr_mz_max <- rx("--max-pr-mz\\s+([\\d.]+)")
  if (pr_mz_min != "n/a" || pr_mz_max != "n/a") info[["Precursor m/z"]] <- paste0(pr_mz_min, "-", pr_mz_max)

  pr_z_min <- rx("--min-pr-charge\\s+(\\d+)"); pr_z_max <- rx("--max-pr-charge\\s+(\\d+)")
  if (pr_z_min != "n/a" || pr_z_max != "n/a") info[["Precursor Charge"]] <- paste0(pr_z_min, "-", pr_z_max)

  fr_min <- rx("--min-fr-mz\\s+([\\d.]+)"); fr_max <- rx("--max-fr-mz\\s+([\\d.]+)")
  if (fr_min != "n/a" || fr_max != "n/a") info[["Fragment m/z"]] <- paste0(fr_min, "-", fr_max)

  info[["Enzyme Cut Sites"]]           <- rx("--cut\\s+(\\S+)")
  info[["Missed Cleavages"]]           <- rx("--missed-cleavages\\s+(\\d+)")
  info[["N-term Met Excision"]]        <- ifelse(grepl("--met-excision", text), "Yes", "No")
  info[["Fixed Modification (Cys)"]]   <- ifelse(grepl("--unimod4", text),
                                            "Carbamidomethylation (Unimod 4)", "None / not set")

  # --var-mod UniMod:35,15.994915,M  (note: "--var-mods N" is the per-peptide maximum)
  vm <- unlist(regmatches(text, gregexpr("--var-mod\\s+\\S+", text, perl = TRUE)))
  if (length(vm) > 0) {
    unimod_names <- c("UNIMOD:35" = "Oxidation", "UNIMOD:1" = "Acetyl",
                      "UNIMOD:21" = "Phospho",   "UNIMOD:121" = "GlyGly",
                      "UNIMOD:7"  = "Deamidation")
    vm_txt <- vapply(sub("--var-mod\\s+", "", vm), function(x) {
      parts <- strsplit(x, ",", fixed = TRUE)[[1]]
      nm    <- unimod_names[toupper(parts[1])]
      site  <- if (length(parts) >= 3) parts[length(parts)] else ""
      if (site == "*n") site <- "protein N-term"
      if (is.na(nm)) x else if (nchar(site) > 0) sprintf("%s (%s)", nm, site) else unname(nm)
    }, character(1))
    info[["Variable Modifications"]] <- paste(unique(vm_txt), collapse = "; ")
  } else {
    info[["Variable Modifications"]] <- "None"
  }
  info[["Max Variable Mods / Peptide"]] <- rx("--var-mods\\s+(\\d+)")
  info[["Cross-run Normalisation"]]     <- ifelse(grepl("--no-norm", text, fixed = TRUE),
                                                  "Disabled (--no-norm)", "Enabled (DIA-NN default)")

  info[["MS2 Mass Accuracy (ppm)"]]    <- rx("--mass-acc\\s+([\\d.]+)")
  info[["MS1 Mass Accuracy (ppm)"]]    <- rx("--mass-acc-ms1\\s+([\\d.]+)")

  n_f <- length(unlist(regmatches(text, gregexpr("--f\\s+\\S+", text, perl = TRUE))))
  if (n_f > 0) info[["Input Raw Files (count)"]] <- n_f

  m_pg <- regmatches(text,
    regexpr("Protein groups with global q-value <= [\\d.]+:\\s*(\\d+)", text, perl = TRUE))
  if (length(m_pg) > 0) {
    info[["Protein Groups (q <= 0.01)"]] <-
      as.integer(sub(".*:\\s*(\\d+)$", "\\1", m_pg, perl = TRUE))
    # Label with the threshold DIA-NN actually reported
    thr <- sub(".*<= ([\\d.]+):.*", "\\1", m_pg, perl = TRUE)
    names(info)[names(info) == "Protein Groups (q <= 0.01)"] <-
      sprintf("Protein Groups (global q <= %s)", thr)
  }

  info
}

load_stats <- function(result_dir) {
  for (fname in c("report.stats.tsv", "report-first-pass.stats.tsv")) {
    p <- file.path(result_dir, fname)
    if (!file.exists(p)) next
    df <- read.delim(p, stringsAsFactors = FALSE, check.names = FALSE)
    if ("File.Name" %in% names(df)) {
      df <- cbind(
        Sample = tools::file_path_sans_ext(basename(df$File.Name)),
        df[, setdiff(names(df), "File.Name"), drop = FALSE]
      )
    }
    return(list(df = df, source = fname))
  }
  list(df = data.frame(), source = "not found")
}

find_pg_matrix <- function(result_dir) {
  p <- file.path(result_dir, "report.pg_matrix.tsv")
  if (file.exists(p)) return(p)
  hits <- list.files(result_dir, pattern = "report\\..*\\.pg_matrix\\.tsv$",
                     full.names = TRUE)
  if (length(hits) > 0) return(hits[1])
  hits <- list.files(result_dir, pattern = "pg_matrix.*\\.tsv$", full.names = TRUE)
  if (length(hits) > 0) return(hits[1])
  NULL
}

# -----------------------------------------------------------------------------
# SUMMARY COMPUTATIONS
# -----------------------------------------------------------------------------

run_quality_summary <- function(stats_df) {
  cols <- c(
    "Precursors.Identified", "Proteins.Identified",
    "FWHM.RT",
    "Median.Mass.Acc.MS1.Corrected", "Median.Mass.Acc.MS2.Corrected",
    "Average.Peptide.Length", "Average.Peptide.Charge",
    "Average.Missed.Tryptic.Cleavages",
    "Normalisation.Instability", "Median.RT.Prediction.Acc"
  )
  avail <- intersect(cols, names(stats_df))
  if (length(avail) == 0) return(data.frame())

  sub <- stats_df[, avail, drop = FALSE]
  data.frame(
    Metric    = avail,
    Mean      = round(colMeans(sub, na.rm = TRUE), 4),
    Median    = round(apply(sub, 2, median, na.rm = TRUE), 4),
    `Std Dev` = round(apply(sub, 2, sd,     na.rm = TRUE), 4),
    Min       = round(apply(sub, 2, function(x) if (all(is.na(x))) NA_real_ else min(x, na.rm = TRUE)), 4),
    Max       = round(apply(sub, 2, function(x) if (all(is.na(x))) NA_real_ else max(x, na.rm = TRUE)), 4),
    `CV (%)`  = round(
      apply(sub, 2, sd, na.rm = TRUE) / colMeans(sub, na.rm = TRUE) * 100, 2),
    check.names = FALSE, row.names = NULL
  )
}

pg_overview <- function(pg_df) {
  sc    <- get_sample_cols(pg_df)
  quant <- pg_df[, sc, drop = FALSE]
  quant[quant == 0] <- NA
  per_sample <- colSums(!is.na(quant))
  any_quant  <- rowSums(!is.na(quant))
  n          <- nrow(pg_df)

  data.frame(
    Metric = c(
      "Total Protein Groups",
      "Total Samples",
      "Mean Proteins Quantified / Sample",
      "Min Proteins Quantified / Sample",
      "Max Proteins Quantified / Sample",
      "Proteins Detected in ALL Samples",
      "Proteins Detected in >= 75% of Samples",
      "Proteins Detected in >= 50% of Samples"
    ),
    Value = c(
      n,
      length(sc),
      round(mean(per_sample), 1),
      min(per_sample),
      max(per_sample),
      sum(any_quant == length(sc)),
      sum(any_quant / length(sc) >= 0.75),
      sum(any_quant / length(sc) >= 0.50)
    ),
    stringsAsFactors = FALSE
  )
}

# -----------------------------------------------------------------------------
# EXCEL STYLE HELPERS
# -----------------------------------------------------------------------------

hs <- function(bg = CLR_DARK_BLUE, fg = CLR_WHITE, bold = TRUE, wrap = TRUE) {
  createStyle(fgFill = bg, fontColour = fg,
              textDecoration = if (bold) "bold" else NULL,
              halign = "center", valign = "center",
              wrapText = wrap, fontSize = 10, fontName = BASE_FONT,
              border = "TopBottomLeftRight", borderColour = bg)
}

write_table <- function(wb, sheet, df, start_row = 1, start_col = 1,
                        hdr_bg = CLR_DARK_BLUE) {
  if (nrow(df) == 0 || ncol(df) == 0) return(start_row)

  writeData(wb, sheet, df,
            startRow    = start_row,
            startCol    = start_col,
            headerStyle = hs(bg = hdr_bg),
            borders     = "surrounding",
            borderStyle = "thin",
            borderColour = CLR_RULE)


  # Alternate row shading (seq(2, 1) would count backwards for a 1-row table)
  if (nrow(df) >= 2) {
    alt <- createStyle(fgFill = CLR_LIGHT_BLUE)
    addStyle(wb, sheet, alt,
             rows = start_row + seq(2, nrow(df), by = 2),
             cols = start_col:(start_col + ncol(df) - 1),
             gridExpand = TRUE, stack = TRUE)
  }

  start_row + nrow(df) + 1
}

add_title <- function(wb, sheet, title, subtitle = "", row = 1) {
  writeData(wb, sheet, title, startRow = row, startCol = 1)
  addStyle(wb, sheet,
           createStyle(fontColour = CLR_DARK_BLUE, textDecoration = "bold",
                       fontSize = 16, fontName = BASE_FONT),
           rows = row, cols = 1)
  setRowHeights(wb, sheet, rows = row, heights = 26)
  # Accent rule under the title
  addStyle(wb, sheet,
           createStyle(border = "bottom", borderColour = CLR_DARK_BLUE, borderStyle = "medium"),
           rows = row, cols = 1:6, gridExpand = TRUE, stack = TRUE)
  if (nchar(subtitle) > 0) {
    writeData(wb, sheet, subtitle, startRow = row + 1, startCol = 1)
    addStyle(wb, sheet,
             createStyle(fontColour = CLR_TEXT_MUTED, fontSize = 9, fontName = BASE_FONT),
             rows = row + 1, cols = 1)
    return(row + 3)
  }
  row + 2
}

# Worksheet in the report style: no gridlines on text sheets, signed footer,
# landscape page setup.
# NOTE: no tabColour - the installed openxlsx writes a duplicate <worksheet>
# tag when a tab colour is set, and Excel then has to repair every sheet.
add_sheet <- function(wb, name, grid = FALSE) {
  addWorksheet(wb, name, gridLines = grid,
               footer = c(sprintf("&8Prepared by %s - %s", REPORT_AUTHOR, REPORT_AFFILIATION),
                          "&8&[Tab]",
                          "&8Page &[Page] of &[Pages]"))
  pageSetup(wb, name, orientation = "landscape", fitToWidth = TRUE)
}

# -----------------------------------------------------------------------------
# RUN STATISTICS GLOSSARY
# -----------------------------------------------------------------------------

STATS_GLOSSARY <- c(
  "Sample"                           = "Sample (raw file) name.",
  "Precursors.Identified"            = "Precursors (peptide + charge state) identified at 1% FDR in this run.",
  "Proteins.Identified"              = "Protein groups identified at 1% FDR in this run (run-level count - see the note on the Summary Statistics sheet).",
  "Total.Quantity"                   = "Sum of all precursor quantities in the run - a rough measure of total signal / amount loaded.",
  "MS1.Signal"                       = "Total MS1 ion signal recorded in the run.",
  "MS2.Signal"                       = "Total MS2 ion signal recorded in the run.",
  "FWHM.Scans"                       = "Median chromatographic peak width (full width at half maximum) in scans. Too few points per peak reduces quantification precision.",
  "FWHM.RT"                          = "Median chromatographic peak width (full width at half maximum) in minutes.",
  "Median.Mass.Acc.MS1"              = "Median MS1 mass error (ppm) before DIA-NN recalibration.",
  "Median.Mass.Acc.MS1.Corrected"    = "Median MS1 mass error (ppm) after recalibration - values close to 0 are good.",
  "Median.Mass.Acc.MS2"              = "Median MS2 mass error (ppm) before DIA-NN recalibration.",
  "Median.Mass.Acc.MS2.Corrected"    = "Median MS2 mass error (ppm) after recalibration - values close to 0 are good.",
  "MS2.Mass.Instability"             = "Spread of the MS2 mass error along the run (ppm). High values suggest calibration drift.",
  "Normalisation.Instability"        = "How much the normalisation factor varies along the gradient. High values can point to spray or loading problems.",
  "Median.RT.Prediction.Acc"         = "Median difference (minutes) between observed and predicted retention time - smaller is better.",
  "Average.Peptide.Length"           = "Average length (amino acids) of identified peptides.",
  "Average.Peptide.Charge"           = "Average charge state of identified precursors.",
  "Average.Missed.Tryptic.Cleavages" = "Average number of missed trypsin cleavages per peptide. High values can indicate incomplete digestion."
)

# -----------------------------------------------------------------------------
# PER-SAMPLE COMPLETENESS
# -----------------------------------------------------------------------------

# Protein groups quantified and missing % per sample (contaminants excluded)
sample_completeness <- function(pg_df, is_cont) {
  sc  <- get_sample_cols(pg_df)
  mat <- as.matrix(pg_df[!is_cont, sc, drop = FALSE])
  mat[mat == 0] <- NA
  n   <- nrow(mat)
  q   <- colSums(!is.na(mat))
  data.frame(
    Sample                        = vapply(sc, shorten_name, character(1), USE.NAMES = FALSE),
    `Protein Groups Quantified`   = as.integer(q),
    `Missing (%)`                 = if (n > 0) round((n - q) / n * 100, 1) else NA_real_,
    check.names = FALSE, stringsAsFactors = FALSE, row.names = NULL
  )
}

# -----------------------------------------------------------------------------
# FIGURES
# -----------------------------------------------------------------------------

# Renders one PNG with base graphics; returns the path or NULL on failure
render_png <- function(width_in, height_in, draw) {
  f <- normalizePath(tempfile(fileext = ".png"), mustWork = FALSE)
  png(f, width = round(width_in * 110), height = round(height_in * 110), res = 110, type = PNG_TYPE)
  ok <- tryCatch({ draw(); TRUE }, error = function(e) {
    cat(sprintf("  [WARN] Figure rendering failed: %s\n", conditionMessage(e))); FALSE
  }, finally = dev.off())
  if (ok && file.exists(f)) f else NULL
}

# Bottom margin (lines) that fits rotated sample labels
label_margin <- function(labels) min(12, max(4.5, max(nchar(labels)) * 0.42 + 1))

# Shared base-graphics theme for all figures
fig_theme <- function(...) {
  par(family = "sans", bty = "l", las = 1, tcl = -0.25, mgp = c(2.6, 0.55, 0),
      col.axis = FIG_TEXT, col.lab = FIG_TEXT, col.main = FIG_DARK, fg = FIG_TEXT,
      font.main = 2, cex.main = 1.05, cex.axis = 0.8, cex.lab = 0.9, ...)
}

# Bar chart with light horizontal gridlines behind the bars
fig_bar <- function(values, labels, col, main, ylab, ylim = NULL) {
  if (is.null(ylim)) ylim <- c(0, max(values, na.rm = TRUE) * 1.08)
  bp <- barplot(values, names.arg = rep("", length(values)), col = col, border = NA,
                main = main, ylab = ylab, ylim = ylim, axes = FALSE, space = 0.25)
  abline(h = axTicks(2), col = FIG_GRID, lwd = 0.8)
  barplot(values, names.arg = rep("", length(values)), col = col, border = NA,
          ylim = ylim, axes = FALSE, space = 0.25, add = TRUE)
  axis(2, at = axTicks(2), col = NA, col.ticks = FIG_TEXT, labels = format(axTicks(2), big.mark = ",", scientific = FALSE))
  axis(1, at = bp, labels = labels, las = 2, cex.axis = 0.7, col = NA, col.ticks = NA)
  invisible(bp)
}

# Adds the Figures sheet. Returns invisibly.
add_figures_sheet <- function(wb, stats_df, pg_df, is_cont) {
  sheet <- "Figures"
  add_sheet(wb, sheet)
  r <- add_title(wb, sheet, "Figures",
                 "Overview plots for this experiment. Contaminant protein groups are excluded from the intensity plots.")
  setColWidths(wb, sheet, cols = 1, widths = 12)

  place <- function(file, caption, note, w, h) {
    if (is.null(file)) return(invisible(NULL))
    writeData(wb, sheet, caption, startRow = r, startCol = 1)
    addStyle(wb, sheet, createStyle(textDecoration = "bold", fontColour = CLR_DARK_BLUE, fontSize = 12),
             rows = r, cols = 1)
    writeData(wb, sheet, note, startRow = r + 1, startCol = 1)
    addStyle(wb, sheet, createStyle(fontColour = "#595959", textDecoration = "italic", fontSize = 9),
             rows = r + 1, cols = 1)
    insertImage(wb, sheet, file, startRow = r + 2, startCol = 1, width = w, height = h, units = "in")
    r <<- r + 2 + ceiling(h * 72 / 15) + 2   # default row = 15 pt
    invisible(NULL)
  }

  # Figure 1 - identifications per sample (Run Statistics)
  if (nrow(stats_df) > 0 && all(c("Precursors.Identified", "Proteins.Identified") %in% names(stats_df))) {
    labs <- as.character(stats_df[[1]])
    w <- max(8, min(16, 4 + nrow(stats_df) * 0.25))
    f <- render_png(w, 5, function() {
      fig_theme(mfrow = c(1, 2), mar = c(label_margin(labs), 5, 2.5, 0.5))
      fig_bar(stats_df$Precursors.Identified, labs, FIG_MAIN, "Precursors identified", "Count")
      fig_bar(stats_df$Proteins.Identified,   labs, FIG_DARK, "Protein groups identified", "Count")
    })
    place(f, "Figure 1 - Identifications per sample",
          "Run-level counts at 1% FDR from DIA-NN (report.stats.tsv). A sample far below the others may indicate a sample or instrument problem.",
          w, 5)
  }

  sc <- get_sample_cols(pg_df)
  if (length(sc) == 0) return(invisible(NULL))
  labs <- vapply(sc, shorten_name, character(1), USE.NAMES = FALSE)
  mat  <- as.matrix(pg_df[!is_cont, sc, drop = FALSE])
  mat[mat == 0] <- NA
  lmat <- log2(mat)
  lmat[!is.finite(lmat)] <- NA
  is_qc <- grepl("qc", labs, ignore.case = TRUE)
  wide  <- max(8, min(16, 4 + length(sc) * 0.25))

  # Figure 2 - intensity distributions
  f <- render_png(wide, 5, function() {
    fig_theme(mar = c(label_margin(labs), 4.5, 2.5, 0.5))
    box_fill <- ifelse(is_qc, adjustcolor(FIG_ACCENT, 0.35), adjustcolor(FIG_MAIN, 0.25))
    boxplot(as.data.frame(lmat), names = rep("", length(labs)), outline = FALSE, axes = FALSE,
            col = box_fill, border = FIG_DARK, medcol = FIG_DARK, whisklty = 1, staplelty = 0,
            main = "Protein group intensity distribution", ylab = "log2 intensity")
    abline(h = axTicks(2), col = FIG_GRID, lwd = 0.8)
    boxplot(as.data.frame(lmat), names = rep("", length(labs)), outline = FALSE, axes = FALSE,
            col = box_fill, border = FIG_DARK, medcol = FIG_DARK, whisklty = 1, staplelty = 0, add = TRUE)
    axis(2, col = NA, col.ticks = FIG_TEXT)
    axis(1, at = seq_along(labs), labels = labs, las = 2, cex.axis = 0.7, col = NA, col.ticks = NA)
    if (any(is_qc)) legend("topright", c("Sample", "QC"), fill = c(adjustcolor(FIG_MAIN, 0.25), adjustcolor(FIG_ACCENT, 0.35)),
                           border = FIG_DARK, bty = "n", cex = 0.8)
  })
  place(f, "Figure 2 - Intensity distribution per sample",
        "log2 protein group intensities (pg_matrix). After normalisation the medians should be similar; a shifted box suggests a loading or normalisation problem.",
        wide, 5)

  # Figure 3 - missing values and data completeness
  n_pg <- nrow(lmat)
  if (n_pg > 0) {
    miss_pct  <- colSums(is.na(lmat)) / n_pg * 100
    det_frac  <- rowSums(!is.na(lmat)) / ncol(lmat)
    thresholds <- seq(0, 1, by = 0.05)
    curve_n   <- vapply(thresholds, function(t) sum(det_frac >= t & det_frac > 0), numeric(1))
    f <- render_png(wide, 5, function() {
      fig_theme(mfrow = c(1, 2), mar = c(label_margin(labs), 4.5, 2.5, 0.5))
      fig_bar(miss_pct, labs, FIG_ACCENT, "Missing values per sample", "Missing protein groups (%)",
              ylim = c(0, max(10, ceiling(max(miss_pct, na.rm = TRUE) / 10) * 10)))
      par(mar = c(label_margin(labs), 5, 2.5, 0.5))
      plot(thresholds * 100, curve_n, type = "n", axes = FALSE,
           xlab = "Quantified in at least X% of samples", ylab = "Protein groups",
           main = "Data completeness")
      abline(h = axTicks(2), v = axTicks(1), col = FIG_GRID, lwd = 0.8)
      lines(thresholds * 100, curve_n, col = FIG_MAIN, lwd = 2)
      points(thresholds * 100, curve_n, pch = 19, cex = 0.7, col = FIG_DARK)
      axis(1, col = NA, col.ticks = FIG_TEXT)
      axis(2, at = axTicks(2), col = NA, col.ticks = FIG_TEXT, labels = format(axTicks(2), big.mark = ",", scientific = FALSE))
    })
    place(f, "Figure 3 - Missing values and data completeness",
          "Left: share of protein groups not quantified in each sample. Right: how many protein groups are quantified in at least X% of samples (useful when choosing a missing-value filter).",
          wide, 5)
  }

  # Figure 4 - sample correlation heatmap
  if (ncol(lmat) >= 2) {
    cm <- suppressWarnings(cor(lmat, use = "pairwise.complete.obs"))
    dimnames(cm) <- list(labs, labs)
    ord <- seq_len(ncol(cm))
    if (ncol(cm) > 2 && !anyNA(cm)) ord <- hclust(as.dist(1 - cm))$order
    cm  <- cm[ord, ord]
    rng <- range(cm, na.rm = TRUE)
    side <- max(6, min(12, 3 + ncol(cm) * 0.22))
    f <- render_png(side + 1.2, side, function() {
      layout(matrix(c(1, 2), 1), widths = c(side, 1.2))
      show_labels <- ncol(cm) <= 90
      lab_cex <- if (ncol(cm) > 40) 0.45 else 0.65
      lm <- if (show_labels) label_margin(colnames(cm)) * (lab_cex / 0.7) + 1 else 1.5
      fig_theme(mar = c(lm, lm, 2.5, 0.5))
      pal <- hcl.colors(64, "Blues 3", rev = TRUE)
      image(seq_len(ncol(cm)), seq_len(nrow(cm)), cm, col = pal, zlim = rng,
            axes = FALSE, xlab = "", ylab = "", main = "Sample correlation (Pearson, log2)")
      if (show_labels) {
        axis(1, at = seq_len(ncol(cm)), labels = colnames(cm), las = 2, cex.axis = lab_cex, tick = FALSE, line = -0.6)
        axis(2, at = seq_len(nrow(cm)), labels = rownames(cm), las = 2, cex.axis = lab_cex, tick = FALSE, line = -0.6)
      }
      par(mar = c(lm, 0.5, 2.5, 3))
      image(1, seq(rng[1], rng[2], length.out = 64), matrix(seq(rng[1], rng[2], length.out = 64), 1),
            col = pal, axes = FALSE, xlab = "", ylab = "")
      axis(4, las = 1, cex.axis = 0.7)
      mtext("r", side = 3, line = 0.5, cex = 0.8)
    })
    place(f, "Figure 4 - Sample correlation",
          sprintf("Pearson correlation of log2 intensities (r = %.3f to %.3f), samples ordered by similarity. Replicates should cluster together; an isolated sample may be an outlier.",
                  rng[1], rng[2]),
          side + 1.2, side)
  }
  invisible(NULL)
}


# -----------------------------------------------------------------------------
# REPORT 1  - QC METRICS
# -----------------------------------------------------------------------------

build_report <- function(project_dir, result_dir, out_path, second_dir = NULL) {
  cat("\nBuilding Analysis_Report.xlsx ...\n")
  wb    <- createWorkbook(creator = REPORT_AUTHOR, title = "Proteomics Analysis Report")
  modifyBaseFont(wb, fontSize = 10, fontName = BASE_FONT)
  pname <- basename(project_dir)
  now   <- format(Sys.time(), "%Y-%m-%d %H:%M")

  # -- Load project metadata (ProjectID, PI) from project_info.json ------------
  read_json_str <- function(path, field) {
    txt <- tryCatch(paste(readLines(path, warn = FALSE, encoding = "UTF-8"), collapse = " "),
                   error = function(e) "")
    m <- regmatches(txt, regexpr(
      sprintf('"%s"\\s*:\\s*"([^"]*)"', field), txt, perl = TRUE))
    if (length(m) == 0) return("")
    sub(sprintf('"%s"\\s*:\\s*"([^"]*)"', field), "\\1", m, perl = TRUE)
  }

  proj_id <- ""
  pi_name <- ""
  for (ip in c(file.path(project_dir, "project_info.json"),
               file.path(dirname(project_dir), "project_info.json"))) {
    if (file.exists(ip)) {
      proj_id <- read_json_str(ip, "ProjectID")
      pi_name <- read_json_str(ip, "PI")
      break
    }
  }

  log_info  <- parse_log(file.path(result_dir, "report.log.txt"))
  qm        <- log_info[["Quantification Method"]]
  is_legacy <- isTRUE(qm == "Legacy")
  qm_label  <- if (qm %in% c("QuantUMS", "MaxLFQ", "Legacy")) qm else "DIA-NN"
  dn_ver    <- log_info[["DIA-NN Version"]]
  fdr_txt   <- sub("\\s.*$", "", log_info[["FDR Threshold (q-value)"]])
  fdr_num   <- suppressWarnings(as.numeric(fdr_txt))
  if (is.na(fdr_num)) { fdr_num <- 0.01; fdr_txt <- "0.01" }
  fdr_pct   <- paste0(format(fdr_num * 100, drop0trailing = TRUE), "%")
  norm_on   <- !grepl("^Disabled", log_info[["Cross-run Normalisation"]])

  subtitle_parts <- c(if (nchar(proj_id) > 0) sprintf("Project ID: %s", proj_id),
                      if (nchar(pi_name)  > 0) sprintf("PI: %s",         pi_name),
                      sprintf("Generated: %s", now),
                      sprintf("Analysis software: DIA-NN%s",
                              if (identical(dn_ver, "n/a")) "" else paste0(" ", sub("\\s.*$", "", dn_ver))))

  # -- Sheet: Project Overview -------------------------------------------------
  cat("  * Project overview & DIA-NN parameters\n")
  add_sheet(wb, "Project Overview")

  r <- add_title(wb, "Project Overview",
                 pname,
                 paste(subtitle_parts, collapse = "   |   "))

  # -- Narrative: how to use this report ----------------------------------------
  writeData(wb, "Project Overview",
            "How to Use This Report", startRow = r, startCol = 1)
  addStyle(wb, "Project Overview",
           createStyle(fontColour = CLR_MID_BLUE, textDecoration = "bold",
                       fgFill = CLR_SECTION, fontSize = 12),
           rows = r, cols = 1:2, gridExpand = TRUE, stack = TRUE)
  mergeCells(wb, "Project Overview", cols = 1:2, rows = r)
  r <- r + 1

  intro <- paste0(
    "This workbook contains the results of a DIA-NN data-independent acquisition (DIA) ",
    "proteomics analysis. The tabs below provide quality-control metrics, run diagnostics, ",
    "and the final protein quantification matrix."
  )
  writeData(wb, "Project Overview", intro, startRow = r, startCol = 1)
  addStyle(wb, "Project Overview",
           createStyle(wrapText = TRUE, fontSize = 10),
           rows = r, cols = 1:2, gridExpand = TRUE, stack = TRUE)
  mergeCells(wb, "Project Overview", cols = 1:2, rows = r)
  setRowHeights(wb, "Project Overview", rows = r, heights = 30)
  r <- r + 2

  sheet_guide <- data.frame(
    Tab = c(
      "1. Project Overview",
      "2. Raw Files",
      "3. Run Statistics",
      "4. Summary Statistics",
      "5. Figures",
      "6. Protein Groups (pg_matrix)   (final result)"
    ),
    Contents = c(
      "DIA-NN analysis parameters, software version, and run metadata.",
      "MS raw file inventory: file names, sizes, and file dates (last modified, close to acquisition time).",
      "Per-sample DIA-NN identification counts, mass accuracy and RT metrics, with a glossary and colour legend.",
      "Cross-sample QC summary, protein group detection overview, missing values per sample, and QC replicate CVs.",
      "Plots: identifications per sample, intensity distributions, missing values / completeness, sample correlation.",
      paste0("Full protein group quantification matrix (", qm_label, " intensities). Primary deliverable for downstream analysis.")
    ),
    stringsAsFactors = FALSE, check.names = FALSE
  )
  r <- write_table(wb, "Project Overview", sheet_guide,
                   start_row = r, hdr_bg = CLR_MID_BLUE) + 1

  writeData(wb, "Project Overview",
            "Key Notes for Data Interpretation", startRow = r, startCol = 1)
  addStyle(wb, "Project Overview",
           createStyle(fontColour = CLR_MID_BLUE, textDecoration = "bold",
                       fgFill = CLR_SECTION, fontSize = 12),
           rows = r, cols = 1:2, gridExpand = TRUE, stack = TRUE)
  mergeCells(wb, "Project Overview", cols = 1:2, rows = r)
  r <- r + 1

  for (note in c(
    paste0("\u2022  FINAL RESULT: The 'Protein Groups (pg_matrix)' tab (Sheet 6) is the ",
           "final quantification output and the recommended starting point for all downstream analysis."),
    paste0("\u2022  LOG2 TRANSFORMATION: Raw intensities are ",
           switch(qm_label, Legacy = "direct (legacy) quantification values", MaxLFQ = "MaxLFQ values",
                  QuantUMS = "QuantUMS values", "DIA-NN quantities"),
           " on a linear scale. ",
           "Log2 transformation is strongly recommended before statistical testing, PCA, heatmaps, or volcano plots."),
    paste0("\u2022  MISSING VALUES: A value of zero or blank means the protein was not detected in that ",
           "sample (below detection threshold or q-value > ", fdr_txt, "). Treat these as NA or apply imputation before downstream analysis. ",
           "Missing values per sample are listed on the Summary Statistics sheet."),
    paste0("\u2022  Q-VALUE FILTER: All reported identifications pass DIA-NN's ", fdr_pct, " FDR filter ",
           "(precursor and protein q-value <= ", fdr_txt, ")."),
    paste0("\u2022  CONTAMINANTS: Protein groups from the contaminant database (e.g. cRAP entries such as keratins and trypsin) ",
           "are counted on the Summary Statistics sheet and left out of the figures. They remain in the pg_matrix; remove them before biological interpretation."),
    if (norm_on)
      paste0("\u2022  NORMALISATION: Intensities are already normalised across runs by DIA-NN",
             if (qm_label == "QuantUMS") " (QuantUMS pipeline)." else ".",
             " Do not apply a second global normalisation without good reason.")
    else
      paste0("\u2022  NORMALISATION: DIA-NN cross-run normalisation was DISABLED for this analysis (--no-norm). ",
             "Intensities are not normalised; consider normalising before comparing samples.")
  )) {
    writeData(wb, "Project Overview", note, startRow = r, startCol = 1)
    addStyle(wb, "Project Overview",
             createStyle(wrapText = TRUE, fontSize = 10),
             rows = r, cols = 1:2, gridExpand = TRUE, stack = TRUE)
    mergeCells(wb, "Project Overview", cols = 1:2, rows = r)
    setRowHeights(wb, "Project Overview", rows = r, heights = 40)
    r <- r + 1
  }
  r <- r + 1  # blank spacer

  disclaimer <- paste0(
    "Report generated automatically by the ", REPORT_AFFILIATION, ". ",
    "For questions, reanalysis requests, or additional statistical support, ",
    "please contact your core facility staff."
  )
  writeData(wb, "Project Overview", disclaimer, startRow = r, startCol = 1)
  addStyle(wb, "Project Overview",
           createStyle(fontColour = "#595959", textDecoration = "italic",
                       wrapText = TRUE, fontSize = 9),
           rows = r, cols = 1:2, gridExpand = TRUE, stack = TRUE)
  mergeCells(wb, "Project Overview", cols = 1:2, rows = r)
  setRowHeights(wb, "Project Overview", rows = r, heights = 35)
  r <- r + 2

  # -- Column descriptions: pg_matrix ------------------------------------------
  writeData(wb, "Project Overview",
            "Column Descriptions: Protein Groups (pg_matrix)", startRow = r, startCol = 1)
  addStyle(wb, "Project Overview",
           createStyle(fontColour = CLR_MID_BLUE, textDecoration = "bold",
                       fgFill = CLR_SECTION, fontSize = 12),
           rows = r, cols = 1:2, gridExpand = TRUE, stack = TRUE)
  mergeCells(wb, "Project Overview", cols = 1:2, rows = r)
  r <- r + 1

  col_desc <- data.frame(
    Column = c(
      "Protein.Group",
      "Protein.Names",
      "Genes",
      "First.Protein.Description",
      "N.Sequences",
      "N.Proteotypic.Sequences",
      "[Sample columns]"
    ),
    Description = c(
      "UniProt accession of the protein group (leading protein). Primary identifier for matching across databases.",
      "Full protein name(s) for all members of the group.",
      "Gene symbol(s) associated with the protein group (as given in the FASTA database).",
      "Functional description of the leading (highest-confidence) protein in the group.",
      "Total number of peptide sequences identified and used for quantification of this protein group.",
      "Number of proteotypic (peptides unique to this protein, not shared with any other) sequences - a measure of identification confidence.",
      paste0(qm_label,
             " protein intensity for each sample (column header = raw file name). Blank = not detected; treat as NA or apply imputation.")
    ),
    stringsAsFactors = FALSE, check.names = FALSE
  )
  r <- write_table(wb, "Project Overview", col_desc,
                   start_row = r, hdr_bg = CLR_MID_BLUE) + 1

  r <- r + 1  # spacer before DIA-NN parameters section

  # -- DIA-NN run parameters ---------------------------------------------------
  kv <- data.frame(
    Parameter = c("Project Folder", "Result Folder",
                  if (!is.null(second_dir)) "Extra Raw Files Folder" else NULL,
                  "Report Generated", "",
                  paste0(rep(" ", nchar("DIA-NN Run Parameters")), collapse = ""),
                  names(log_info)),
    Value = c(project_dir, result_dir,
              if (!is.null(second_dir)) second_dir else NULL,
              format(Sys.time(), "%Y-%m-%d %H:%M:%S"), "",
              " - DIA-NN Run Parameters  -",
              unlist(log_info)),
    stringsAsFactors = FALSE
  )

  writeData(wb, "Project Overview", kv,
            startRow = r, startCol = 1, colNames = FALSE)
  addStyle(wb, "Project Overview",
           createStyle(textDecoration = "bold", fontColour = CLR_DARK_BLUE),
           rows = r:(r + nrow(kv) - 1), cols = 1,
           gridExpand = TRUE, stack = TRUE)
  divider_row <- r - 1L + which(kv$Value == " - DIA-NN Run Parameters  -")[1]
  addStyle(wb, "Project Overview",
           createStyle(fgFill = CLR_SECTION, textDecoration = "bold",
                       fontColour = CLR_MID_BLUE),
           rows = divider_row, cols = 1:2,
           gridExpand = TRUE, stack = TRUE)

  setColWidths(wb, "Project Overview", cols = 1:2, widths = c(45, 85))

  # -- Signature (bottom of the sheet) ----------------------------------------------
  rs  <- r + nrow(kv) + 2
  sig <- data.frame(Field = c("Prepared by", "Affiliation", "Date"),
                    Value = c(REPORT_AUTHOR, REPORT_AFFILIATION, format(Sys.Date(), "%d %B %Y")),
                    stringsAsFactors = FALSE)
  writeData(wb, "Project Overview", sig, startRow = rs, startCol = 1, colNames = FALSE)
  addStyle(wb, "Project Overview", createStyle(textDecoration = "bold", fontColour = CLR_DARK_BLUE),
           rows = rs:(rs + 2), cols = 1, gridExpand = TRUE, stack = TRUE)
  addStyle(wb, "Project Overview", createStyle(border = "top", borderColour = CLR_DARK_BLUE),
           rows = rs, cols = 1:2, gridExpand = TRUE, stack = TRUE)

  # -- Sheet: Raw Files --------------------------------------------------------
  cat("  * Raw file inventory\n")
  add_sheet(wb, "Raw Files")
  raw_df <- collect_raw_files(project_dir, second_dir)

  if (nrow(raw_df) > 0) {
    n_raw_cols <- ncol(raw_df)
    total_mb <- sum(raw_df$`Size (MB)`)
    r2 <- add_title(wb, "Raw Files",
                    "Raw Mass Spectrometry Files",
                    sprintf("Count: %d   |   Total size: %s   |   Average size: %s",
                            nrow(raw_df),
                            bytes_to_human(total_mb * 1024^2),
                            bytes_to_human(mean(raw_df$`Size (MB)`) * 1024^2)))
    end_r2 <- write_table(wb, "Raw Files", raw_df, start_row = r2)

    total_row <- setNames(as.list(rep("", n_raw_cols)), names(raw_df))
    total_row[["File Name"]] <- "TOTAL"
    total_row[["Size (MB)"]] <- sprintf("%.2f MB  (%s)", total_mb, bytes_to_human(total_mb * 1024^2))
    writeData(wb, "Raw Files", as.data.frame(total_row, check.names = FALSE),
              startRow = end_r2 + 1, startCol = 1, colNames = FALSE)
    addStyle(wb, "Raw Files",
             createStyle(textDecoration = "bold"),
             rows = end_r2 + 1, cols = seq_len(n_raw_cols), gridExpand = TRUE, stack = TRUE)

    raw_note_row <- end_r2 + 3
    writeData(wb, "Raw Files",
              paste0("Note: Raw data files (.raw) are not included in the standard delivery. ",
                     "Please contact your core facility staff if you require access to the original raw files."),
              startRow = raw_note_row, startCol = 1)
    addStyle(wb, "Raw Files",
             createStyle(fontColour = "#595959", textDecoration = "italic",
                         wrapText = TRUE, fontSize = 9),
             rows = raw_note_row, cols = seq_len(n_raw_cols), gridExpand = TRUE, stack = TRUE)
    mergeCells(wb, "Raw Files", cols = seq_len(n_raw_cols), rows = raw_note_row)
    setRowHeights(wb, "Raw Files", rows = raw_note_row, heights = 28)
  } else {
    n_raw_cols <- 4L
    no_raw_msg <- if (!is.null(second_dir))
      "No raw files found in project folder or extra raw files folder." else
      "No raw files found in project directory."
    writeData(wb, "Raw Files", no_raw_msg, startRow = 1, startCol = 1)
  }
  raw_col_widths <- if ("Source Folder" %in% names(raw_df)) c(30, 25, 12, 10, 18) else c(30, 12, 10, 18)
  setColWidths(wb, "Raw Files", cols = seq_len(n_raw_cols), widths = raw_col_widths)

  # -- Sheet: Run Statistics ---------------------------------------------------
  cat("  * Per-sample run statistics\n")
  add_sheet(wb, "Run Statistics", grid = TRUE)
  stats_res <- load_stats(result_dir)
  stats_df  <- stats_res$df
  stats_src <- stats_res$source

  if (nrow(stats_df) > 0) {
    num_idx <- sapply(stats_df, is.numeric)
    stats_df[, num_idx] <- round(stats_df[, num_idx], 4)

    r3 <- add_title(wb, "Run Statistics",
                    "DIA-NN Per-Sample Run Statistics",
                    sprintf("Source: %s   |   Samples: %d", stats_src, nrow(stats_df)))
    end_r3 <- write_table(wb, "Run Statistics", stats_df, start_row = r3)

    rr <- end_r3 + 1
    if ("Proteins.Identified" %in% names(stats_df)) {
      ci    <- which(names(stats_df) == "Proteins.Identified")
      med_p <- median(stats_df$Proteins.Identified, na.rm = TRUE)
      for (i in seq_len(nrow(stats_df))) {
        val <- stats_df$Proteins.Identified[i]
        bg  <- if (!is.na(val) && val >= med_p * 0.90) CLR_GOOD else
               if (!is.na(val) && val >= med_p * 0.70) CLR_WARN else CLR_BAD
        addStyle(wb, "Run Statistics",
                 createStyle(fgFill = bg),
                 rows = r3 + i, cols = ci, stack = TRUE)
      }

      # Colour legend
      writeData(wb, "Run Statistics", "Colour legend for Proteins.Identified", startRow = rr, startCol = 1)
      addStyle(wb, "Run Statistics", createStyle(textDecoration = "bold", fontColour = CLR_DARK_BLUE),
               rows = rr, cols = 1)
      legend_rows <- list(
        list(CLR_GOOD, sprintf("Green  - at least 90%% of the median (%s protein groups)", format(med_p, big.mark = ","))),
        list(CLR_WARN, "Amber  - 70-90% of the median"),
        list(CLR_BAD,  "Red    - below 70% of the median: check this sample")
      )
      for (k in seq_along(legend_rows)) {
        addStyle(wb, "Run Statistics", createStyle(fgFill = legend_rows[[k]][[1]]), rows = rr + k, cols = 1)
        writeData(wb, "Run Statistics", legend_rows[[k]][[2]], startRow = rr + k, startCol = 2)
      }
      writeData(wb, "Run Statistics",
                "Samples are compared with the median of all samples in this report. If the report mixes different sample types, differences may be biological rather than technical.",
                startRow = rr + 4, startCol = 2)
      addStyle(wb, "Run Statistics", createStyle(fontColour = "#595959", textDecoration = "italic", fontSize = 9),
               rows = rr + 4, cols = 2)
      rr <- rr + 6
    }

    # Glossary of the columns present
    gl_cols <- intersect(names(stats_df), names(STATS_GLOSSARY))
    if (length(gl_cols) > 0) {
      writeData(wb, "Run Statistics", "Column glossary", startRow = rr, startCol = 1)
      addStyle(wb, "Run Statistics", createStyle(textDecoration = "bold", fontColour = CLR_DARK_BLUE),
               rows = rr, cols = 1)
      gl <- data.frame(Column = gl_cols, Meaning = unname(STATS_GLOSSARY[gl_cols]),
                       stringsAsFactors = FALSE)
      writeData(wb, "Run Statistics", gl, startRow = rr + 1, startCol = 1, colNames = FALSE)
      addStyle(wb, "Run Statistics", createStyle(textDecoration = "bold"),
               rows = (rr + 1):(rr + nrow(gl)), cols = 1, gridExpand = TRUE, stack = TRUE)
    }
    col_w <- pmin(pmax(nchar(names(stats_df)) + 2, 10), 22)
    setColWidths(wb, "Run Statistics",
                 cols = seq_along(stats_df), widths = col_w)
  } else {
    writeData(wb, "Run Statistics", "report.stats.tsv not found in Result folder.",
              startRow = 1, startCol = 1)
  }

  # -- Sheet: Summary Statistics -----------------------------------------------
  cat("  * Summary statistics\n")
  add_sheet(wb, "Summary Statistics")
  r4 <- add_title(wb, "Summary Statistics", "Run Quality Summary  - All Samples")

  if (nrow(stats_df) > 0) {
    summ <- run_quality_summary(stats_df)
    if (nrow(summ) > 0)
      r4 <- write_table(wb, "Summary Statistics", summ, start_row = r4) + 2
  }

  pg_path <- find_pg_matrix(result_dir)
  if (!is.null(pg_path)) {
    pg_raw  <- read.delim(pg_path, stringsAsFactors = FALSE, check.names = FALSE)
    is_cont <- flag_contaminants(pg_raw, log_info[["Contaminant Exclusion Tag"]])

    writeData(wb, "Summary Statistics",
              "Protein Group Matrix Overview", startRow = r4, startCol = 1)
    addStyle(wb, "Summary Statistics",
             createStyle(fgFill = CLR_SECTION, fontColour = CLR_MID_BLUE,
                         textDecoration = "bold"),
             rows = r4, cols = 1, stack = TRUE)
    r4 <- r4 + 1
    pg_ov <- pg_overview(pg_raw)
    sc_ov <- get_sample_cols(pg_raw)
    m_ov  <- as.matrix(pg_raw[, sc_ov, drop = FALSE]); m_ov[m_ov == 0] <- NA
    pg_ov <- rbind(pg_ov, data.frame(
      Metric = c("Contaminant Protein Groups (flagged)", "Missing Values, all samples (%)"),
      Value  = c(sum(is_cont), if (length(m_ov) > 0) round(mean(is.na(m_ov)) * 100, 1) else NA),
      stringsAsFactors = FALSE))
    r4 <- write_table(wb, "Summary Statistics", pg_ov,
                      start_row = r4, hdr_bg = CLR_MID_BLUE)

    # Why these counts differ from Proteins.Identified on Run Statistics
    count_note <- paste0(
      "Note: 'Proteins Quantified / Sample' counts protein groups with a non-zero intensity in the final ",
      "pg_matrix. The matrix is filtered at the experiment-wide (global) protein q-value and can include quantities ",
      "supported by evidence from other runs (match-between-runs). 'Proteins.Identified' on the Run Statistics ",
      "sheet is DIA-NN's count of protein groups identified within that run alone, so the two numbers are expected to differ.")
    writeData(wb, "Summary Statistics", count_note, startRow = r4, startCol = 1)
    addStyle(wb, "Summary Statistics",
             createStyle(fontColour = "#595959", textDecoration = "italic", fontSize = 9,
                         wrapText = TRUE, valign = "top"),
             rows = r4, cols = 1:7, gridExpand = TRUE, stack = TRUE)
    mergeCells(wb, "Summary Statistics", cols = 1:7, rows = r4)
    setRowHeights(wb, "Summary Statistics", rows = r4, heights = 50)
    r4 <- r4 + 2

    # -- Missing values per sample -----------------------------------------------
    writeData(wb, "Summary Statistics",
              "Missing Values per Sample (contaminants excluded)", startRow = r4, startCol = 1)
    addStyle(wb, "Summary Statistics",
             createStyle(fgFill = CLR_SECTION, fontColour = CLR_MID_BLUE, textDecoration = "bold"),
             rows = r4, cols = 1, stack = TRUE)
    r4 <- r4 + 1
    r4 <- write_table(wb, "Summary Statistics", sample_completeness(pg_raw, is_cont),
                      start_row = r4, hdr_bg = CLR_MID_BLUE) + 1

    # -- QC Sample Metrics ------------------------------------------------------
    # QC samples = sample columns whose short name (not the full path) contains "qc"
    sc_all  <- get_sample_cols(pg_raw)
    qc_cols <- sc_all[grepl("qc", vapply(sc_all, shorten_name, character(1)), ignore.case = TRUE)]
    n_qc    <- length(qc_cols)

    if (n_qc >= 2) {
      qc_short      <- vapply(qc_cols, shorten_name, character(1), USE.NAMES = FALSE)
      # Strip a separated replicate number (QC100_1 -> QC100); otherwise strip
      # only a short 1-2 digit suffix (QC1 -> QC) so QC100 / QC200 stay distinct
      qc_groups     <- ifelse(grepl("[_-]\\d+$", qc_short),
                              sub("[_-]\\d+$", "", qc_short),
                              sub("(?<!\\d)\\d{1,2}$", "", qc_short, perl = TRUE))
      qc_groups     <- ifelse(qc_groups == "", qc_short, qc_groups)
      unique_groups <- unique(qc_groups)
      n_groups      <- length(unique_groups)
      grp_palette   <- c(FIG_MAIN, FIG_ACCENT, "#548235", "#7F6000", "#7030A0")

      writeData(wb, "Summary Statistics",
                "QC Sample Metrics", startRow = r4, startCol = 1)
      addStyle(wb, "Summary Statistics",
               createStyle(fgFill = CLR_SECTION, fontColour = CLR_MID_BLUE,
                           textDecoration = "bold"),
               rows = r4, cols = 1, stack = TRUE)
      r4 <- r4 + 1

      # Warning: any group with fewer than 3 replicates
      small_groups <- unique_groups[
        sapply(unique_groups, function(g) sum(qc_groups == g) < 3)
      ]
      if (length(small_groups) > 0) {
        writeData(wb, "Summary Statistics",
                  paste0("Warning: fewer than 3 replicates in group(s): ",
                         paste(small_groups, collapse = ", "),
                         ". CV% requires >= 3 replicates for reliable interpretation."),
                  startRow = r4, startCol = 1)
        addStyle(wb, "Summary Statistics",
                 createStyle(fontColour = "#9C5700", fgFill = CLR_WARN,
                             textDecoration = "italic", fontSize = 10, wrapText = TRUE),
                 rows = r4, cols = 1:2, gridExpand = TRUE, stack = TRUE)
        mergeCells(wb, "Summary Statistics", cols = 1:2, rows = r4)
        setRowHeights(wb, "Summary Statistics", rows = r4, heights = 28)
        r4 <- r4 + 2
      }

      qc_table_row    <- r4   # anchor for image alignment
      first_grp_end   <- NULL
      group_data      <- list()

      for (gi in seq_along(unique_groups)) {
        grp      <- unique_groups[gi]
        grp_idx  <- which(qc_groups == grp)
        grp_cols <- qc_cols[grp_idx]
        grp_mat  <- as.matrix(pg_raw[, grp_cols, drop = FALSE])
        grp_mat[grp_mat == 0] <- NA

        per_samp <- colSums(!is.na(grp_mat))
        in_all   <- sum(rowSums(!is.na(grp_mat)) == length(grp_cols))
        cv_prot  <- apply(grp_mat, 1, function(x) {
          x <- x[!is.na(x)]
          if (length(x) < 2 || mean(x) == 0) return(NA_real_)
          sd(x) / mean(x) * 100
        })
        med_cv <- round(median(cv_prot, na.rm = TRUE), 2)
        grp_color <- grp_palette[((gi - 1L) %% length(grp_palette)) + 1L]

        group_data[[grp]] <- list(
          mat    = grp_mat,
          short  = qc_short[grp_idx],
          cv     = cv_prot,
          med_cv = med_cv,
          color  = grp_color
        )

        # Sub-header
        writeData(wb, "Summary Statistics", grp, startRow = r4, startCol = 1)
        addStyle(wb, "Summary Statistics",
                 createStyle(textDecoration = "bold", fontColour = CLR_DARK_BLUE,
                             fgFill = CLR_LIGHT_BLUE),
                 rows = r4, cols = 1:2, gridExpand = TRUE, stack = TRUE)
        r4 <- r4 + 1

        grp_summ <- data.frame(
          Metric = c(
            "QC Samples",
            "Mean Proteins / QC Sample",
            "Min Proteins / QC Sample",
            "Max Proteins / QC Sample",
            "Proteins in ALL QC Samples",
            "Median CV% Across QC Samples"
          ),
          Value = c(
            length(grp_cols),
            round(mean(per_samp), 1),
            min(per_samp),
            max(per_samp),
            in_all,
            if (is.na(med_cv)) "n/a" else as.character(med_cv)
          ),
          stringsAsFactors = FALSE
        )
        r4 <- write_table(wb, "Summary Statistics", grp_summ,
                          start_row = r4, hdr_bg = CLR_MID_BLUE)
        if (gi == 1L) first_grp_end <- r4
        if (gi < n_groups) r4 <- r4 + 1   # spacer between groups
      }

      qc_row_ht_pts <- 30L
      n_img_rows    <- if (n_groups == 1L) r4 - qc_table_row else first_grp_end - qc_table_row
      qc_img_height <- max(4.5, n_img_rows * qc_row_ht_pts / 72)
      setRowHeights(wb, "Summary Statistics",
                    rows    = qc_table_row:(r4 - 1L),
                    heights = qc_row_ht_pts)
      r4 <- r4 + 1   # final spacer

      # Plot
      qc_plot_file <- normalizePath(tempfile(fileext = ".png"), mustWork = FALSE)
      cat(sprintf("  * Rendering QC plot -> %s\n", qc_plot_file))

      px_w <- 1000L
      px_h <- round(px_w * qc_img_height / 8)
      png(qc_plot_file, width = px_w, height = px_h, res = 110, type = PNG_TYPE)
      tryCatch({
        all_short <- unlist(lapply(unique_groups, function(g) group_data[[g]]$short))
        fig_theme(mfrow = c(1, 2), mar = c(label_margin(all_short), 4, 2.5, 0.5), oma = c(0, 0, 1.2, 0))

        # Left: boxplots colored by group
        box_data  <- list()
        box_names <- character(0)
        box_cols  <- character(0)
        for (grp in unique_groups) {
          gd      <- group_data[[grp]]
          log2mat <- log2(gd$mat)
          log2mat[!is.finite(log2mat)] <- NA
          for (j in seq_len(ncol(log2mat))) {
            box_data[[length(box_data) + 1L]] <- log2mat[, j]
            box_names <- c(box_names, gd$short[j])
            box_cols  <- c(box_cols,  gd$color)
          }
        }
        # Group colours are explained by the legend in the CV panel
        boxplot(box_data, names = rep("", length(box_data)), outline = FALSE, axes = FALSE,
                col = adjustcolor(box_cols, 0.45), border = FIG_DARK, medcol = FIG_DARK,
                whisklty = 1, staplelty = 0,
                main = "Intensity distribution (log2)", ylab = "log2 intensity")
        abline(h = axTicks(2), col = FIG_GRID, lwd = 0.8)
        boxplot(box_data, names = rep("", length(box_data)), outline = FALSE, axes = FALSE,
                col = adjustcolor(box_cols, 0.45), border = FIG_DARK, medcol = FIG_DARK,
                whisklty = 1, staplelty = 0, add = TRUE)
        axis(2, col = NA, col.ticks = FIG_TEXT)
        axis(1, at = seq_along(box_names), labels = box_names, las = 2, cex.axis = 0.7,
             col = NA, col.ticks = NA)

        # Right: overlaid CV% histograms (density scale) + median lines
        all_cv <- unlist(lapply(group_data, function(g) g$cv[is.finite(g$cv)]))
        cv_xlim <- if (length(all_cv) > 0) range(all_cv) else c(0, 100)
        # Base histogram on the first group that has finite CVs
        cv_grps   <- unique_groups[sapply(unique_groups, function(g)
                       any(is.finite(group_data[[g]]$cv)))]
        first_grp <- if (length(cv_grps) > 0) cv_grps[1] else unique_groups[1]
        cv_vals   <- group_data[[first_grp]]$cv
        cv_vals   <- cv_vals[is.finite(cv_vals)]
        if (length(cv_vals) == 0) {
          plot.new()
          title(main = "Per-Protein CV% Distribution",
                xlab = "CV% across QC samples", ylab = "Density")
          text(0.5, 0.5, "Insufficient overlapping detections\nto compute CV",
               cex = 0.85, col = "gray50", adj = c(0.5, 0.5))
        } else {
        hist(cv_vals, breaks = 40, freq = FALSE,
             col  = adjustcolor(group_data[[first_grp]]$color, alpha.f = 0.45),
             border = NA, xlim = cv_xlim,
             main = "Per-Protein CV% Distribution",
             xlab = "CV% across QC samples", ylab = "Density")
        if (n_groups > 1) {
          for (grp in setdiff(unique_groups, first_grp)) {
            cv_vals <- group_data[[grp]]$cv
            cv_vals <- cv_vals[is.finite(cv_vals)]
            if (length(cv_vals) == 0) next
            hist(cv_vals, breaks = 40, freq = FALSE, add = TRUE,
                 col = adjustcolor(group_data[[grp]]$color, alpha.f = 0.45),
                 border = NA)
          }
        }
        for (grp in unique_groups)
          abline(v = group_data[[grp]]$med_cv,
                 col = group_data[[grp]]$color, lwd = 2, lty = 2)
        legend("topright",
               legend = sapply(unique_groups,
                               function(g) sprintf("%s: %.1f%%", g, group_data[[g]]$med_cv)),
               col = sapply(unique_groups, function(g) group_data[[g]]$color),
               lwd = 2, lty = 2, bty = "n", cex = 0.8)
        }  # end else (cv_vals non-empty)

        mtext("QC Sample Diagnostics", outer = TRUE, cex = 1.1, font = 2)
      }, error = function(e) {
        cat(sprintf("  [WARN] QC plot rendering failed: %s\n", conditionMessage(e)))
      }, finally = {
        dev.off()
      })

      if (file.exists(qc_plot_file)) {
        insertImage(wb, "Summary Statistics", qc_plot_file,
                    startRow = qc_table_row, startCol = 4,
                    width = 8, height = qc_img_height, units = "in")
        cat("  * QC plot inserted into Summary Statistics\n")
      } else {
        cat("  [WARN] QC plot file not found, skipping image insert\n")
      }
    }
  }

  setColWidths(wb, "Summary Statistics", cols = 1:7,
               widths = c(40, 12, 12, 12, 12, 12, 10))

  # -- Sheet: Figures -----------------------------------------------------------
  cat("  * Figures\n")
  if (!is.null(pg_path)) {
    add_figures_sheet(wb, stats_df, pg_raw, is_cont)
  } else if (nrow(stats_df) > 0) {
    add_figures_sheet(wb, stats_df, data.frame(), logical(0))
  }

  # -- Sheet: Protein Groups (pg_matrix) - direct copy -------------------------
  # pg_path and pg_raw already loaded above for Summary Statistics; reuse them
  if (!is.null(pg_path)) {
    # Exact copy of the DIA-NN file (same columns, values and order - no added
    # columns or highlighting); only the sample headers are reduced from the
    # full raw-file path to the file name
    cat("  * Writing protein group matrix (exact copy)\n")
    pg_df  <- pg_raw
    sc     <- get_sample_cols(pg_df)
    names(pg_df)[match(sc, names(pg_df))] <- vapply(sc, shorten_name, character(1), USE.NAMES = FALSE)
    sc     <- get_sample_cols(pg_df)
    add_sheet(wb, "Protein Groups (pg_matrix)", grid = TRUE)
    r6 <- add_title(wb, "Protein Groups (pg_matrix)",
                    sprintf("Protein Group Quantification Matrix (%s)", qm_label))
    writeData(wb, "Protein Groups (pg_matrix)", pg_df,
              startRow    = r6,
              startCol    = 1,
              headerStyle = hs(bg = CLR_DARK_BLUE),
              keepNA      = FALSE)
    col_w <- c(Protein.Group = 25, Protein.Ids = 25, Protein.Names = 20, Genes = 12,
               First.Protein.Description = 45, N.Sequences = 12,
               N.Proteotypic.Sequences = 22)
    widths <- vapply(names(pg_df), function(n) if (n %in% names(col_w)) col_w[[n]] else 18, numeric(1))
    setColWidths(wb, "Protein Groups (pg_matrix)",
                 cols = seq_along(pg_df), widths = unname(widths))
    freezePane(wb, "Protein Groups (pg_matrix)", firstActiveRow = r6 + 1, firstActiveCol = 2)
  } else {
    cat("  [WARN] Protein group matrix not found - sheet skipped.\n")
  }

  # openxlsx only warns when the target is locked (e.g. open in Excel) and the
  # old file stays in place - check first and fail loudly instead
  if (file.exists(out_path)) {
    con <- tryCatch(suppressWarnings(file(out_path, open = "ab")), error = function(e) NULL)
    if (is.null(con))
      stop(sprintf("Cannot write %s - the file is open in another program (close it in Excel and run again).", out_path),
           call. = FALSE)
    close(con)
  }
  saveWorkbook(wb, out_path, overwrite = TRUE)
  cat(sprintf("  OK  Saved -> %s\n", out_path))
}

# (build_final_results removed - pg_matrix is now Sheet 6 of Analysis_Report.xlsx)

# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------

if (!dir.exists(project_dir))
  stop(sprintf("[ERROR] Project directory not found: %s", project_dir))

if (!dir.exists(result_dir)) {
  cat(sprintf("[WARN] '%s' subfolder not found  - scanning for alternatives ...\n",
              result_dir_name))
  candidates <- list.dirs(project_dir, recursive = FALSE, full.names = TRUE)
  candidates <- candidates[grepl("[Rr]esult", basename(candidates))]
  if (length(candidates) == 0) stop("[ERROR] No result directory found. Aborting.")
  result_dir <- candidates[1]
  cat(sprintf("       Using: %s\n", result_dir))
}

if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)

out_path <- file.path(output_dir, "Analysis_Report.xlsx")

build_report(project_dir, result_dir, out_path, second_dir)

cat(sprintf("\n%s\n", sep))
cat(" Done!\n")
cat(sprintf("   Analysis Report  ->  %s\n", out_path))
cat(sprintf("%s\n\n", sep))
