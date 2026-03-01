#!/usr/bin/env Rscript
# =============================================================================
# Proteomics Service Report Generator  - DIA-NN Output
# =============================================================================
# Generates one Excel file from a DIA-NN project folder:
#
#   Analysis_Report.xlsx  - QC metrics (4 sheets) + direct copy of report.pg_matrix.tsv
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
RAW_EXTENSIONS  <- c(".raw")

CLR_DARK_BLUE  <- "#1F4E79"
CLR_MID_BLUE   <- "#2E75B6"
CLR_LIGHT_BLUE <- "#D9E1F2"
CLR_WHITE      <- "#FFFFFF"
CLR_GOOD       <- "#C6EFCE"
CLR_WARN       <- "#FFEB9C"
CLR_BAD        <- "#FFC7CE"
CLR_SECTION    <- "#EEF3FA"

PG_META <- c("Protein.Group", "Protein.Names", "Genes",
             "First.Protein.Description", "N.Sequences",
             "N.Proteotypic.Sequences")

# --- Arguments ----------------------------------------------------------------
args <- commandArgs(trailingOnly = TRUE)

project_dir_arg <- if (length(args) >= 1) args[1] else getwd()
result_dir_name <- if (length(args) >= 2) args[2] else RESULT_DIR_NAME
output_dir_arg  <- if (length(args) >= 3) args[3] else project_dir_arg

project_dir <- normalizePath(project_dir_arg, winslash = "/", mustWork = FALSE)
result_dir  <- file.path(project_dir, result_dir_name)
output_dir  <- normalizePath(output_dir_arg, winslash = "/", mustWork = FALSE)

sep <- strrep("=", 62)
cat(sprintf("\n%s\n", sep))
cat(" Proteomics Report Generator  |  DIA-NN Output\n")
cat(sprintf("%s\n", sep))
cat(sprintf(" Project : %s\n", project_dir))
cat(sprintf(" Results : %s\n", result_dir))
cat(sprintf(" Output  : %s\n", output_dir))
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

get_sample_cols <- function(df) setdiff(names(df), PG_META)

shorten_colnames <- function(df) {
  names(df) <- sapply(names(df), function(n) {
    if (grepl("[/\\\\]", n)) tools::file_path_sans_ext(basename(n)) else n
  })
  df
}

# -----------------------------------------------------------------------------
# DATA COLLECTION
# -----------------------------------------------------------------------------

collect_raw_files <- function(project_dir) {
  files <- character(0)
  for (ext in RAW_EXTENSIONS) {
    files <- c(files,
               list.files(project_dir,
                          pattern    = paste0("\\", ext, "$"),
                          full.names = TRUE,
                          ignore.case = TRUE))
  }
  if (length(files) == 0) return(data.frame())

  rows <- lapply(sort(files), function(f) {
    st    <- file.info(f)
    quant <- paste0(f, ".quant")
    data.frame(
      `File Name`       = basename(f),
      `Size (MB)`       = round(st$size / 1024^2, 2),
      `Size (GB)`       = round(st$size / 1024^3, 3),
      `Modified`        = format(st$mtime, "%Y-%m-%d %H:%M"),
      `Has .quant File` = ifelse(file.exists(quant), "Yes", "No"),
      check.names = FALSE, stringsAsFactors = FALSE
    )
  })
  do.call(rbind, rows)
}

parse_log <- function(log_path) {
  info <- list(
    # --- System ----------------------------------------------------------
    "DIA-NN Version"                = "n/a",
    "Compile Date"                  = "n/a",
    "Analysis Date"                 = "n/a",
    "CPU"                           = "n/a",
    "Threads"                       = "n/a",
    # --- Analysis settings -----------------------------------------------
    "FDR Threshold (q-value)"       = "n/a",
    "Quantification Method"         = "n/a",
    "Match-Between-Runs (MBR)"      = "n/a",
    "Reuse Existing .quant Files"   = "n/a",
    "Output Matrices"               = "n/a",
    "Generate Spectral Library"     = "n/a",
    # --- Library / database ----------------------------------------------
    "Spectral Library"              = "n/a",
    "Output Library"                = "n/a",
    "FASTA Database(s)"             = "n/a",
    "Contaminant Exclusion Tag"     = "n/a",
    # --- Peptide settings ------------------------------------------------
    "Enzyme Cut Sites"              = "n/a",
    "Missed Cleavages"              = "n/a",
    "N-term Met Excision"           = "n/a",
    "Fixed Modification (Cys)"      = "n/a",
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
  info[["Compile Date"]]               <- rx("Compiled on (.+?)\\n")
  info[["Analysis Date"]]              <- rx("Current date and time:\\s*(.+?)\\n")
  info[["CPU"]]                        <- rx("CPU:\\s+(.+?)\\n")
  info[["Threads"]]                    <- rx("--threads\\s+(\\d+)")

  info[["FDR Threshold (q-value)"]]    <- rx("--qvalue\\s+([\\d.]+)")
  major_ver <- suppressWarnings(as.integer(sub("^(\\d+)\\..*", "\\1", ver_str)))
  info[["Quantification Method"]]      <- if (!is.na(major_ver) && major_ver >= 2)
                                            "QuantUMS (DIA-NN 2.x default)" else "MaxLFQ"
  info[["Match-Between-Runs (MBR)"]]   <- ifelse(grepl("--reanalyse",    text), "Yes", "No")
  info[["Reuse Existing .quant Files"]]<- ifelse(grepl("--use-quant",    text), "Yes", "No")
  info[["Output Matrices"]]            <- ifelse(grepl("--matrices",     text), "Yes", "No")
  info[["Generate Spectral Library"]]  <- ifelse(grepl("--gen-spec-lib", text), "Yes", "No")

  m_lib <- regmatches(text, regexpr("--lib\\s+(\\S+)", text, perl = TRUE))
  if (length(m_lib) > 0)
    info[["Spectral Library"]] <- basename(sub("--lib\\s+(\\S+)", "\\1", m_lib, perl = TRUE))

  m_outlib <- regmatches(text, regexpr("--out-lib\\s+(\\S+)", text, perl = TRUE))
  if (length(m_outlib) > 0)
    info[["Output Library"]] <- basename(sub("--out-lib\\s+(\\S+)", "\\1", m_outlib, perl = TRUE))

  fastas <- unlist(regmatches(text, gregexpr("(?<=--fasta\\s)\\S+", text, perl = TRUE)))
  if (length(fastas) > 0)
    info[["FASTA Database(s)"]] <- paste(basename(fastas), collapse = "; ")

  info[["Contaminant Exclusion Tag"]]  <- rx("--cont-quant-exclude\\s+(\\S+)")

  info[["Enzyme Cut Sites"]]           <- rx("--cut\\s+(\\S+)")
  info[["Missed Cleavages"]]           <- rx("--missed-cleavages\\s+(\\d+)")
  info[["N-term Met Excision"]]        <- ifelse(grepl("--met-excision", text), "Yes", "No")
  info[["Fixed Modification (Cys)"]]   <- ifelse(grepl("--unimod4", text),
                                            "Carbamidomethylation (Unimod 4)", "None / not set")

  info[["MS2 Mass Accuracy (ppm)"]]    <- rx("--mass-acc\\s+([\\d.]+)")
  info[["MS1 Mass Accuracy (ppm)"]]    <- rx("--mass-acc-ms1\\s+([\\d.]+)")

  n_f <- length(unlist(regmatches(text, gregexpr("--f\\s+\\S+", text, perl = TRUE))))
  if (n_f > 0) info[["Input Raw Files (count)"]] <- n_f

  m_pg <- regmatches(text,
    regexpr("Protein groups with global q-value <= [\\d.]+:\\s*(\\d+)", text, perl = TRUE))
  if (length(m_pg) > 0)
    info[["Protein Groups (q <= 0.01)"]] <-
      as.integer(sub(".*:\\s*(\\d+)$", "\\1", m_pg, perl = TRUE))

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
    Min       = round(apply(sub, 2, min,    na.rm = TRUE), 4),
    Max       = round(apply(sub, 2, max,    na.rm = TRUE), 4),
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
              wrapText = wrap, fontSize = 11)
}

write_table <- function(wb, sheet, df, start_row = 1, start_col = 1,
                        hdr_bg = CLR_DARK_BLUE) {
  if (nrow(df) == 0 || ncol(df) == 0) return(start_row)

  writeData(wb, sheet, df,
            startRow    = start_row,
            startCol    = start_col,
            headerStyle = hs(bg = hdr_bg),
            borders     = "surrounding",
            borderStyle = "thin")


  alt <- createStyle(fgFill = CLR_LIGHT_BLUE)
  for (i in seq(2, nrow(df), by = 2)) {
    addStyle(wb, sheet, alt,
             rows = start_row + i,
             cols = start_col:(start_col + ncol(df) - 1),
             gridExpand = TRUE, stack = TRUE)
  }

  start_row + nrow(df) + 1
}

add_title <- function(wb, sheet, title, subtitle = "", row = 1) {
  writeData(wb, sheet, title, startRow = row, startCol = 1)
  addStyle(wb, sheet,
           createStyle(fontColour = CLR_DARK_BLUE, textDecoration = "bold",
                       fontSize = 13),
           rows = row, cols = 1)
  if (nchar(subtitle) > 0) {
    writeData(wb, sheet, subtitle, startRow = row + 1, startCol = 1)
    addStyle(wb, sheet,
             createStyle(fontColour = "#595959", textDecoration = "italic",
                         fontSize = 10),
             rows = row + 1, cols = 1)
    return(row + 3)
  }
  row + 2
}

# -----------------------------------------------------------------------------
# FILE DESCRIPTION LOOKUP
# -----------------------------------------------------------------------------

file_description <- function(fname) {
  if (grepl("pg_matrix", fname, fixed = TRUE))
    return("Protein group quantification matrix (QuantUMS)  - FINAL RESULT")

  lut <- c(
    "report.parquet"                      = "Main DIA-NN output  - all precursors/proteins (binary columnar format)",
    "report.log.txt"                      = "DIA-NN execution log  - parameters, diagnostics, run summary",
    "report.manifest.txt"                 = "JSON manifest listing output files and their purpose",
    "report.stats.tsv"                    = "Per-sample statistics: identifications, mass accuracy, RT, charge, etc.",
    "report.pr_matrix.tsv"                = "Precursor-level quantification matrix",
    "report.gg_matrix.tsv"                = "Gene group quantification matrix (MaxLFQ/QuantUMS)",
    "report.unique_genes_matrix.tsv"      = "Proteotypic-gene quantification matrix",
    "report.protein_description.tsv"      = "Protein annotations and descriptions",
    "report-first-pass.parquet"           = "Intermediate first-pass DIA-NN output",
    "report-first-pass.pr_matrix.tsv"     = "First-pass precursor matrix",
    "report-first-pass.stats.tsv"         = "First-pass per-sample statistics",
    "report-first-pass.manifest.txt"      = "First-pass manifest",
    "report-lib.parquet"                  = "Refined spectral library generated from this experiment",
    "report-lib.parquet.skyline.speclib"  = "Spectral library exported for Skyline",
    "report_runs.pdf"                     = "PDF: Per-run QC visualisations",
    "report_trends.pdf"                   = "PDF: Cross-run trend plots"
  )
  d <- lut[fname]
  if (is.na(d)) "DIA-NN output file" else unname(d)
}

# -----------------------------------------------------------------------------
# REPORT 1  - QC METRICS
# -----------------------------------------------------------------------------

build_report <- function(project_dir, result_dir, out_path) {
  cat("\nBuilding Analysis_Report.xlsx ...\n")
  wb    <- createWorkbook(creator = "Kittinun Leetanaporn")
  pname <- basename(project_dir)
  now   <- format(Sys.time(), "%Y-%m-%d %H:%M")

  # -- Sheet: Project Overview -------------------------------------------------
  cat("  * Project overview & DIA-NN parameters\n")
  addWorksheet(wb, "Project Overview")
  log_info <- parse_log(file.path(result_dir, "report.log.txt"))

  r <- add_title(wb, "Project Overview",
                 pname,
                 sprintf("Proteomics Analysis Report   |   Generated: %s   |   Instrument software: DIA-NN", now))

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
      "1 \u2014 Project Overview",
      "2 \u2014 Raw Files",
      "3 \u2014 Run Statistics",
      "4 \u2014 Summary Statistics",
      "5 \u2014 Protein Groups (pg_matrix)   \u2605 FINAL RESULT"
    ),
    Contents = c(
      "DIA-NN analysis parameters, software version, and run metadata.",
      "MS raw file inventory: file sizes and .quant pairing status.",
      "Per-sample DIA-NN identification counts, mass accuracy, RT metrics, and run-quality flags.",
      "Cross-sample means, medians, and CVs for key QC metrics; protein group detection overview.",
      "Full protein group quantification matrix (QuantUMS intensities). Primary deliverable for downstream analysis."
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
    paste0("\u2022  FINAL RESULT \u2014 The \u2018Protein Groups (pg_matrix)\u2019 tab (Sheet 5) is the ",
           "final quantification output and the recommended starting point for all downstream analysis."),
    paste0("\u2022  LOG2 TRANSFORMATION \u2014 Raw intensities are QuantUMS values on a linear scale. ",
           "Log2 transformation is strongly recommended before statistical testing, PCA, heatmaps, or volcano plots."),
    paste0("\u2022  MISSING VALUES \u2014 A value of zero or blank means the protein was not detected in that ",
           "sample (below detection threshold or q-value > 0.01). Treat these as NA or apply imputation before downstream analysis."),
    paste0("\u2022  Q-VALUE FILTER \u2014 All reported identifications pass DIA-NN\u2019s 1% FDR filter ",
           "(precursor and protein q-value \u2264 0.01) unless a different threshold is listed in the run parameters below."),
    paste0("\u26A0  NORMALISATION \u2014 Intensities are already normalised by DIA-NN at the precursor level (QuantUMS pipeline).")
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
    "Report generated automatically by the Proteomics Core Facility. ",
    "For questions, reanalysis requests, or additional statistical support, ",
    "please contact your service representative."
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
      "HGNC gene symbol(s) associated with the protein group.",
      "Functional description of the leading (highest-confidence) protein in the group.",
      "Total number of peptide sequences identified and used for quantification of this protein group.",
      "Number of proteotypic (peptides unique to this protein, not shared with any other) sequences - a measure of identification confidence.",
      "QuantUMS protein intensity for each sample (column header = sample name). Zero or blank = not detected; treat as NA or apply imputation."
    ),
    stringsAsFactors = FALSE, check.names = FALSE
  )
  r <- write_table(wb, "Project Overview", col_desc,
                   start_row = r, hdr_bg = CLR_MID_BLUE) + 1

  r <- r + 1  # spacer before DIA-NN parameters section

  # -- DIA-NN run parameters ---------------------------------------------------
  kv <- data.frame(
    Parameter = c("Project Folder", "Result Folder", "Report Generated", "",
                  paste0(rep(" ", nchar("DIA-NN Run Parameters")), collapse = ""),
                  names(log_info)),
    Value = c(project_dir, result_dir,
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
  divider_row <- r + 4
  addStyle(wb, "Project Overview",
           createStyle(fgFill = CLR_SECTION, textDecoration = "bold",
                       fontColour = CLR_MID_BLUE),
           rows = divider_row, cols = 1:2,
           gridExpand = TRUE, stack = TRUE)

  setColWidths(wb, "Project Overview", cols = 1:2, widths = c(45, 85))

  # -- Sheet: Raw Files --------------------------------------------------------
  cat("  * Raw file inventory\n")
  addWorksheet(wb, "Raw Files")
  raw_df <- collect_raw_files(project_dir)

  if (nrow(raw_df) > 0) {
    total_mb <- sum(raw_df$`Size (MB)`)
    r2 <- add_title(wb, "Raw Files",
                    "Raw Mass Spectrometry Files",
                    sprintf("Count: %d   |   Total size: %s   |   Average size: %s",
                            nrow(raw_df),
                            bytes_to_human(total_mb * 1024^2),
                            bytes_to_human(mean(raw_df$`Size (MB)`) * 1024^2)))
    end_r2 <- write_table(wb, "Raw Files", raw_df, start_row = r2)

    quant_ok <- sum(raw_df$`Has .quant File` == "Yes")
    writeData(wb, "Raw Files",
              data.frame(
                A = "TOTAL",
                B = sprintf("%.2f MB  (%s)", total_mb, bytes_to_human(total_mb * 1024^2)),
                C = "", D = "",
                E = sprintf("%d / %d paired .quant files", quant_ok, nrow(raw_df))
              ),
              startRow = end_r2 + 1, startCol = 1, colNames = FALSE)
    addStyle(wb, "Raw Files",
             createStyle(textDecoration = "bold"),
             rows = end_r2 + 1, cols = 1:5, gridExpand = TRUE, stack = TRUE)

    raw_note_row <- end_r2 + 3
    writeData(wb, "Raw Files",
              paste0("\u2139  Raw data files (.raw) are not included in the standard delivery. ",
                     "Please contact your service representative if you require access to the original raw files."),
              startRow = raw_note_row, startCol = 1)
    addStyle(wb, "Raw Files",
             createStyle(fontColour = "#595959", textDecoration = "italic",
                         wrapText = TRUE, fontSize = 9),
             rows = raw_note_row, cols = 1:5, gridExpand = TRUE, stack = TRUE)
    mergeCells(wb, "Raw Files", cols = 1:5, rows = raw_note_row)
    setRowHeights(wb, "Raw Files", rows = raw_note_row, heights = 28)
  } else {
    writeData(wb, "Raw Files", "No raw files found in project directory.",
              startRow = 1, startCol = 1)
  }
  setColWidths(wb, "Raw Files", cols = 1:5, widths = c(30, 12, 10, 18, 20))

  # -- Sheet: Run Statistics ---------------------------------------------------
  cat("  * Per-sample run statistics\n")
  addWorksheet(wb, "Run Statistics")
  stats_res <- load_stats(result_dir)
  stats_df  <- stats_res$df
  stats_src <- stats_res$source

  if (nrow(stats_df) > 0) {
    num_idx <- sapply(stats_df, is.numeric)
    stats_df[, num_idx] <- round(stats_df[, num_idx], 4)

    r3 <- add_title(wb, "Run Statistics",
                    "DIA-NN Per-Sample Run Statistics",
                    sprintf("Source: %s   |   Samples: %d", stats_src, nrow(stats_df)))
    write_table(wb, "Run Statistics", stats_df, start_row = r3)

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
  addWorksheet(wb, "Summary Statistics")
  r4 <- add_title(wb, "Summary Statistics", "Run Quality Summary  - All Samples")

  if (nrow(stats_df) > 0) {
    summ <- run_quality_summary(stats_df)
    if (nrow(summ) > 0)
      r4 <- write_table(wb, "Summary Statistics", summ, start_row = r4) + 2
  }

  pg_path <- find_pg_matrix(result_dir)
  if (!is.null(pg_path)) {
    pg_raw <- read.delim(pg_path, stringsAsFactors = FALSE, check.names = FALSE)
    writeData(wb, "Summary Statistics",
              "Protein Group Matrix Overview", startRow = r4, startCol = 1)
    addStyle(wb, "Summary Statistics",
             createStyle(fgFill = CLR_SECTION, fontColour = CLR_MID_BLUE,
                         textDecoration = "bold"),
             rows = r4, cols = 1, stack = TRUE)
    r4 <- r4 + 1
    write_table(wb, "Summary Statistics", pg_overview(pg_raw),
                start_row = r4, hdr_bg = CLR_MID_BLUE)
  }

  setColWidths(wb, "Summary Statistics", cols = 1:7,
               widths = c(40, 12, 12, 12, 12, 12, 10))

  # -- Sheet: Protein Groups (pg_matrix) — direct copy -------------------------
  pg_path <- find_pg_matrix(result_dir)
  if (!is.null(pg_path)) {
    cat("  * Writing protein group matrix (direct copy)\n")
    pg_raw <- read.delim(pg_path, stringsAsFactors = FALSE, check.names = FALSE)
    pg_df  <- shorten_colnames(pg_raw)
    sc     <- get_sample_cols(pg_df)
    addWorksheet(wb, "Protein Groups (pg_matrix)")
    r6 <- add_title(wb, "Protein Groups (pg_matrix)",
                    "Protein Group Quantification Matrix (QuantUMS)",
                    sprintf("Source: %s   |   Protein groups: %d   |   Samples: %d   |   DIA-NN q-value <= 0.01",
                            basename(pg_path), nrow(pg_df), length(sc)))
    write_table(wb, "Protein Groups (pg_matrix)", pg_df, start_row = r6)
    meta_w   <- c(25, 20, 12, 45, 12, 22)[seq_along(intersect(PG_META, names(pg_df)))]
    sample_w <- rep(18, length(sc))
    setColWidths(wb, "Protein Groups (pg_matrix)",
                 cols = seq_along(pg_df), widths = c(meta_w, sample_w))
  } else {
    cat("  [WARN] Protein group matrix not found — sheet skipped.\n")
  }

  saveWorkbook(wb, out_path, overwrite = TRUE)
  cat(sprintf("  OK  Saved -> %s\n", out_path))
}

# (build_final_results removed — pg_matrix is now Sheet 6 of Analysis_Report.xlsx)

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

build_report(project_dir, result_dir, out_path)

cat(sprintf("\n%s\n", sep))
cat(" Done!\n")
cat(sprintf("   Analysis Report  ->  %s\n", out_path))
cat(sprintf("%s\n\n", sep))
