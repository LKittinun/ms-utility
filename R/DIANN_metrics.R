# --- Package bootstrap --------------------------------------------------------
for (pkg in c("diann", "tidyverse", "GGally", "viridis", "gghighlight", "ggridges", "cowplot")) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    message(sprintf("Installing missing package: %s", pkg))
    install.packages(pkg, repos = "https://cloud.r-project.org", quiet = TRUE)
  }
}
if (!requireNamespace("ComplexHeatmap", quietly = TRUE)) {
  message("Installing missing package: ComplexHeatmap")
  if (!requireNamespace("BiocManager", quietly = TRUE))
    install.packages("BiocManager", repos = "https://cloud.r-project.org", quiet = TRUE)
  BiocManager::install("ComplexHeatmap", ask = FALSE, quiet = TRUE)
}

suppressPackageStartupMessages({
library(diann, quietly = T)
library(tidyverse, quietly = T)
library(GGally, quietly = T)
library(viridis, quietly = T)
library(gghighlight, quietly = T)
library(ggridges, quietly = T)
library(cowplot, quietly = T)
library(ComplexHeatmap, quietly = T)
})

suppressMessages( options(warn=-1) )
args <- commandArgs(trailingOnly = T)

folder_path <- args
  #"C:/DIA-NN/1.9/Result/4_Rattus_ziptip_3k_06_07_2024/VF"

exp_name <- tail(unlist(strsplit(folder_path, split = "/")), 1)

setwd(folder_path)

if (!dir.exists("./result")) {dir.create("./result")}

# -- Load DIA-NN precursor report ----------------------------------------------
# DIA-NN 1.x outputs report.tsv; DIA-NN 2.x outputs report.parquet.
# If the folder contains only report.parquet (and no report.tsv), the arrow
# package is used to read it  - the column structure is identical.
if (file.exists("report.tsv")) {
  message("Loading DIA-NN report (TSV  - DIA-NN 1.x) ...")
  df <- diann_load("report.tsv") |>
    mutate(File.Name = gsub(".*[/\\\\](.*)\\.raw$", "\\1", File.Name))

} else if (file.exists("report.parquet")) {
  message("Loading DIA-NN report (Parquet  - DIA-NN 2.x) ...")
  if (!requireNamespace("arrow", quietly = TRUE)) {
    message("Installing 'arrow' package for parquet support ...")
    install.packages("arrow", repos = "https://cloud.r-project.org", quiet = TRUE)
  }
  df <- arrow::read_parquet("report.parquet") |> as.data.frame() |>
    mutate(File.Name = Run)

} else {
  stop(paste0(
    "No DIA-NN report found in: ", folder_path, "\n",
    "Expected report.tsv (DIA-NN 1.x) or report.parquet (DIA-NN 2.x).\n",
    "For DIA-NN 2.x, point to the Result subfolder, e.g.:\n",
    "  Z:/Proteomics/MyProject/Result"
  ))
}

# DIA-NN 2.x removes First.Protein.Description from the main report;
# load from report.protein_description.tsv (Description column) when absent.
if (!("First.Protein.Description" %in% names(df))) {
  if (file.exists("report.protein_description.tsv")) {
    prot_desc <- read_tsv("report.protein_description.tsv", show_col_types = FALSE) |>
      select(Protein.Group = Protein.Id, First.Protein.Description = Description)
    df <- df |> left_join(prot_desc, by = "Protein.Group")
  } else {
    df$First.Protein.Description <- df$Protein.Names
  }
}

df <- df |>
  arrange(match(File.Name, stringr::str_sort(File.Name, numeric = T)))

df_name <- df |> 
  select(Protein.Group, Genes, Protein.Names, First.Protein.Description) |> 
  distinct()

# -- Protein group matrix ------------------------------------------------------
# DIA-NN 2.x writes report.pg_matrix.tsv; load it directly (faster and avoids
# recomputing LFQ from the precursor report via diann_maxlfq).
# Fall back to diann_maxlfq() for DIA-NN 1.x runs that have no pg_matrix file.
pg_matrix_file <- list.files(".", pattern = "^report.*\\.pg_matrix\\.tsv$")[1]

if (!is.na(pg_matrix_file)) {
  message("Loading protein group matrix from ", pg_matrix_file, " ...")
  pg_raw <- read_tsv(pg_matrix_file, show_col_types = FALSE)
  meta_cols <- c("Protein.Group", "Protein.Ids", "First.Protein.Description", "Genes", "Protein.Names")
  sample_cols <- setdiff(names(pg_raw), meta_cols)
  pg_lfq_nontrunc <- pg_raw |>
    select(Protein.Group,
           any_of(c("Genes", "Protein.Names", "First.Protein.Description")),
           all_of(sample_cols)) |>
    arrange(Protein.Group)
  message("Done!")
} else {
  message("Generating protein group matrix via diann_maxlfq() ...")
  pg_lfq_nontrunc <- df |> filter(Q.Value <= 0.01 & PG.Q.Value <= 0.01) |>
    diann_maxlfq(group.header="Protein.Group", id.header = "Precursor.Id",
                 quantity.header = "Precursor.Normalised") |>
    as.data.frame() |>
    rownames_to_column("Protein.Group") |>
    left_join(df_name, by = "Protein.Group") |>
    relocate(Genes, Protein.Names, First.Protein.Description, .after = Protein.Group) |>
    arrange(Protein.Group)
  message("Done!")
}

pg_lfq <- pg_lfq_nontrunc
colnames(pg_lfq)[-c(1:4)] <- str_trunc(colnames(pg_lfq)[-c(1:4)], width = 20, "left")

pg_log2 <- pg_lfq |> 
  select(-Protein.Names, -Genes, -First.Protein.Description) |> 
  mutate(across(where(is.numeric), log2)) 

df_name_pep <-  df |> 
  select(Stripped.Sequence, Protein.Group, Genes, Protein.Names, First.Protein.Description) |> 
  distinct()

message("Generating peptide matrix...")

pep_normalized_nontrunc <- diann_matrix(df, id.header="Stripped.Sequence", quantity.header = "Precursor.Normalised",
                                        q=0.01, proteotypic.only = T, pg.q = 0.01) |> 
  as.data.frame() |> 
  rownames_to_column("Peptides") |> 
  left_join(df_name_pep, by = c("Peptides" = "Stripped.Sequence")) |> 
  relocate(Protein.Group, Genes, Protein.Names, First.Protein.Description, .after = Peptides) |> 
  arrange(Protein.Group)

message("Done!")

pep_normalized <- pep_normalized_nontrunc
colnames(pep_normalized)[-c(1:5)] <- str_trunc(colnames(pep_normalized)[-c(1:5)], width = 20, "left")

pep_log2 <- pep_normalized |> 
  select(-Protein.Names, -Protein.Group, -Genes, -First.Protein.Description) |> 
  mutate(across(where(is.numeric), log2)) 

## Total proteins identified
message(" ")
sink(paste0("result/summary_", exp_name,".txt"))
  
cat(paste0("Total samples: ", ncol(pg_lfq[,-c(1:4), drop = F])), "\n")
cat(paste0("Total peptides identified: ", nrow(pep_normalized)), "\n")
cat(paste0("Total proteins identified: ", nrow(pg_lfq)), "\n")

diann_save(pg_lfq_nontrunc, file = paste0("result/pg_lfq_",exp_name,".tsv"))
diann_save(pep_normalized_nontrunc, file = paste0("result/pep_normalized_",exp_name,".tsv"))

tot_col <- ncol(pg_lfq[,-c(1:4), drop = F]) 

p_dim <- dplyr::case_when(
  tot_col <= 10 ~ 10,
  tot_col > 10 & tot_col <= 20 ~ tot_col,
  tot_col > 20 & tot_col <= 30 ~ tot_col,
  tot_col > 30 ~ 30
  )
 
if(tot_col >1){

    pg_lfq_name <- pg_lfq |> column_to_rownames("Protein.Group")
    prot_half <- rowSums(!is.na(pg_lfq_name[,-c(1:3)])) >= ncol(pg_lfq_name[,-c(1:3)])*0.5
    prot_half_n <- sum(prot_half)
    
    cat(paste0("Proteins that are quantified in > 50% of experiments: ", prot_half_n, "\n"))
    
    missval <- map(pg_lfq[,-c(1:4), drop = F], ~round(sum(!is.na(.x))/length(.x),2)) |> unlist()
    
    if (sum(missval < 0.5) != 0){ 
      
      cat("Samples with missing value >= 50%:", mapply(paste0, " ", names(missval[missval < 0.5]), "(", 
      (1-missval[missval < 0.5])*100, "%)"), "\n")
      
      cat("Samples with missing value >= 75%: ", mapply(paste0, " ", names(missval[missval < 0.25]), "(", 
      (1-missval[missval < 0.25])*100, "%)"), "\n")
      }
      }
      
    sink()
    cat(readLines(paste0("result/summary_",exp_name,".txt")), sep = "\n")
    
    ## Plot
    message("Generating missing value plot...")
    miss_val_plot <- naniar::vis_miss(pg_lfq[,-c(1:4), drop = F], sort_miss = T, cluster = T, warn_large_data = F)
    
    ggsave(miss_val_plot, filename = "result/plot/hm/miss_val.tiff", height = max(7,p_dim-5), width = max(7,p_dim-5), create.dir = T)
    
    miss_val_case_summary <- naniar::miss_var_summary(pg_lfq_nontrunc[,-c(1:4), drop = F], order = F) |> 
      mutate(tot_n = ifelse(n_miss == 0, nrow(pg_lfq), n_miss*(100/pct_miss))) |> 
      select(variable, tot_n, n_miss, pct_miss)
    miss_val_prot_summary <- pg_lfq_nontrunc[,-c(2:4), drop = F] |> column_to_rownames("Protein.Group") |> 
      t() |> as.data.frame() |> 
      naniar::miss_var_summary(order = F) |> 
      mutate(tot_n = ifelse(n_miss == 0, tot_col, n_miss*(100/pct_miss))) |> 
      left_join(df_name, by = c("variable" = "Protein.Group")) |> 
      arrange(pct_miss) |> 
      select(variable, Protein.Names, First.Protein.Description, tot_n, n_miss, pct_miss)


  write_delim(miss_val_case_summary, file= "result/sample_summary.tsv", delim = "\t")
  write_delim(miss_val_prot_summary, file= "result/prot_summary.tsv", delim = "\t")
  
  message("Done!")

    if(tot_col > 1){
      
      message("Generating protein intensity heatmap ...")
      #BUG
      miss_0 <-  miss_val_prot_summary[miss_val_prot_summary$pct_miss==0,]$variable
      miss_25 <- miss_val_prot_summary[miss_val_prot_summary$pct_miss<=25,]$variable
      miss_50 <- miss_val_prot_summary[miss_val_prot_summary$pct_miss<=50,]$variable
      miss_100 <- miss_val_prot_summary$variable
      #batch_df <- data.frame(row.names = colnames(pg_log2[,-1]), "Run" = 1:ncol(pg_log2[,-1]))
      ht_opt$message <- F

      batch_df <- HeatmapAnnotation(Run = anno_simple(1:ncol(pg_log2[,-1]),  pch = as.character(1:ncol(pg_log2[,-1]) ) ))
      walk2(list(miss_0, miss_25, miss_50, miss_100 ), c(100,75,50,0), 
      \(x,y) {
        tiff(paste0("result/plot/hm/prot_hm_", y, "_", exp_name, ".tiff"), height = max(7,p_dim-20), width = max(7,p_dim-20), units="in",res=300)
        p <- pg_log2 |> 
        as.data.frame() |> 
        dplyr::filter(Protein.Group %in% x )|> 
        column_to_rownames("Protein.Group") |> 
        t() |> scale() |> t() |> as.data.frame() %>%
        mutate(across(where(is.numeric), ~replace_na(.x,0))) %>%
          ComplexHeatmap::Heatmap(column_title = paste0("Protein with at least ", y , "% completeness ", "(", nrow(.) ,")"), name = "Z-score\nintensity",
                  show_row_names = F, top_annotation = batch_df, column_names_rot = 45)
       draw(p)
       dev.off()
      })
  
      message("Done!")
      
      message("Generating pairwise protein correlation plot ...")
      pg_pair_data <- pg_log2 |> column_to_rownames("Protein.Group")
      # Top 100 proteins by mean log2 intensity. Use base-R pairs() instead of
      # ggpairs() — ggpairs builds N^2 ggplot objects which is very slow even
      # with few data points; pairs() renders everything in one device call.
      top100 <- order(rowMeans(pg_pair_data, na.rm = TRUE), decreasing = TRUE)[1:min(100, nrow(pg_pair_data))]
      pg_pair_data <- pg_pair_data[top100, ]
      if (!dir.exists("result/plot")) dir.create("result/plot", recursive = TRUE)
      tiff(paste0("result/plot/proteins_pair_", exp_name, ".tiff"),
           width = p_dim, height = p_dim, units = "in", res = 150, compression = "lzw")
      pairs(pg_pair_data, pch = 16, cex = 0.6, col = rgb(0, 0, 0, 0.4), gap = 0.2)
      dev.off()
      message("Done!")
      
      message("Generating proteins correlation heatmap ...")
      proteins_hm <- pg_log2 |> 
        column_to_rownames("Protein.Group") |> 
          ggcorr(low = viridis(3)[3], mid = viridis(3)[2], high = viridis(3)[1])

        ggsave(proteins_hm, filename = paste0("result/plot/proteins_hm_", exp_name, ".tiff"), width = p_dim, height = p_dim, limitsize = F, create.dir = T)
        message("Done!")
        
        message("Generating CV histogram ...")
      # Protein
        protein_hist <- pg_lfq |> 
          select(-Protein.Names, -Genes, -First.Protein.Description) |>
          column_to_rownames("Protein.Group") |> 
          apply(1, \(x) sd(x, na.rm = T)*100/mean(x, na.rm = T))
        
        protein_cv_20 <- data.frame(protein_hist = protein_hist) |> 
          mutate(cv_cat = case_when(protein_hist <= 10 ~ "<10%", protein_hist >20 ~ ">20%", is.na(protein_hist) ~ NA, .default = "10-20%")) |> 
          mutate(cv_cat = fct_relevel(cv_cat, "<10%", "10-20%", ">20%")) |> 
          mutate(group = as.numeric(cv_cat)) |> 
          filter(!is.na(cv_cat)) |> 
          add_count(cv_cat) |> 
          mutate(cv_count = paste0(cv_cat, " (n = ", n, ")")) |> 
          mutate(cv_count = fct_reorder(cv_count, group))
      
      protein_cv_hist <- suppressMessages(ggplot(protein_cv_20, aes(x = protein_hist, y = cv_count, fill = cv_cat, height = after_stat(count))) + 
          geom_density_ridges(stat = "binline", alpha = 0.5, scale = 1, binwidth = 5)+
          geom_vline(xintercept =  10, linetype = "dashed") +
          geom_vline(xintercept =  20, linetype = "dashed") +
          scale_y_discrete(expand = expansion(add = c(0,1.2))) +
         labs(title = "Protein CV histogram", x = "%CV", y = "Count", fill = "Group")+
  #          caption = bquote(bold("Note:")~"CV is only applicable for technical replicates evaluation")) +
          scale_fill_hue(direction = -1, drop = F)+
          theme_bw() +
          theme(plot.caption = element_text(hjust=0))
            )
      # Peptide
      peptide_hist <- pep_normalized |> 
        select(-Protein.Group,-Genes,-Protein.Names, -First.Protein.Description) |> 
        column_to_rownames("Peptides") |> 
        apply(1, \(x) sd(x, na.rm = T)*100/mean(x, na.rm = T))
                
      peptide_cv_20 <- data.frame(peptide_hist = peptide_hist) |> 
        mutate(cv_cat = case_when(peptide_hist <= 10 ~ "<10%", peptide_hist >20 ~ ">20%", is.na(peptide_hist) ~ NA, .default = "10-20%")) |> 
        mutate(cv_cat = fct_relevel(cv_cat, "<10%", "10-20%", ">20%")) |>         
        filter(!is.na(cv_cat)) |> 
        mutate(group = as.numeric(cv_cat)) |> 
        add_count(cv_cat) |> 
        mutate(cv_count = paste0(cv_cat, " (n = ", n, ")")) |>
        mutate(cv_count = fct_reorder(cv_count, group))
      
      peptide_cv_hist <- suppressMessages(ggplot(peptide_cv_20, aes(x = peptide_hist, y = cv_count, fill = cv_cat, height = after_stat(count))) + 
#          geom_histogram(position = "identity", alpha = 0.5, binwidth = 5) + 
        geom_density_ridges(stat = "binline", alpha = 0.5, scale = 1, binwidth = 5)+
        geom_vline(xintercept =  10, linetype = "dashed") +
        geom_vline(xintercept =  20, linetype = "dashed") +
        scale_y_discrete(expand = expansion(add = c(0,1.2))) +
        labs(title = "Peptide CV histogram", x = "%CV", y = "Count", fill = "Group",
          caption = bquote(bold("Note:")~"CV is only applicable for technical replicates evaluation")) +
        scale_fill_hue(direction = -1, drop = F)+
        theme_bw() +
        theme(plot.caption = element_text(hjust=0))
          )
    
      cv_hist_p <- plot_grid(peptide_cv_hist, protein_cv_hist, ncol = 2)
      
      ggsave(cv_hist_p, filename = paste0("result/plot/cv_hist_", exp_name, ".tiff"), height = 7, width = 16, create.dir = T)
      message("Done!")
    } else {
  message("Total sample = 1, statistical plots were not generated.")
}
  
message("Generating intensity plot...")

mean_intensity_df <- pg_log2 |> 
  as.data.frame() |> 
  column_to_rownames("Protein.Group") |> colMeans(na.rm = T) 

global_mean <- mean(mean_intensity_df)
global_sd <- sd(mean_intensity_df)
mean_intens_p <- ggplot(NULL, aes(x = 1:ncol(pg_log2[,-1]), y = mean_intensity_df)) +
  geom_point() +
  geom_hline(yintercept = c(global_mean, global_mean-global_sd,global_mean+global_sd,
                            global_mean -2*global_sd, global_mean+2*global_sd), linetype = "dashed",
                          col = c("black", "blue", "blue", "darkred", "darkred"), alpha = 0.5) +
  labs(Title = "Intensity plot", x = "Run", y = "Mean log2 sample intensity") +
  theme_classic()
ggsave(paste0("result/plot/mean_intensity_plot_",exp_name,".tiff"), width = max(7,p_dim-7), height = 5,create.dir = T)

pg_normalized_intensity <- pg_log2 |> 
  select(-Protein.Group) |> 
  pivot_longer(everything(), names_to = "Sample", values_to = "Intensity") |> 
  ggplot(aes(x = Sample, y = Intensity, fill = Sample)) + 
  geom_boxplot(width = 0.5) +
  theme_bw() +
  labs(title = "Protein normalized intensity", y = "Log2 intensity") +
  scale_x_discrete(labels = \(x) str_trunc(x, 20, side = "left")) +
  guides(fill = "none") +
  theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust=1))

pep_normalized_intensity <- pep_log2 |> 
  select(-Peptides) |> 
  pivot_longer(everything(), names_to = "Sample", values_to = "Intensity") |> 
  ggplot(aes(x = Sample, y = Intensity, fill = Sample)) + 
  geom_boxplot(width = 0.5) +
  theme_bw() +
  scale_x_discrete(labels = \(x) str_trunc(x, 20, side = "left")) +
  labs(title = "Peptide normalized intensity", y = "Log2 intensity") +
  guides(fill = "none") +
  theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust=1))

normalized_intensity_p <- plot_grid(pep_normalized_intensity, pg_normalized_intensity, nrow = 2)

ggsave(paste0("result/plot/normalized_intensity_",exp_name,".tiff"), normalized_intensity_p, width = p_dim-5, height = 10, create.dir = T)
message("Done!")

message("Generating precursor intensity vs retention time plot ...")

precursor_int_df <- df |> 
  filter(Q.Value <= 0.01 & Proteotypic == 1) |> 
  select(Modified.Sequence, Precursor.Normalised, RT.Start, RT.Stop) |> 
  summarize(across(everything(), .fns = mean),.by = Modified.Sequence) |> 
  mutate(RT.Width = RT.Stop - RT.Start) |> 
  mutate(Precursor.Normalised = log2(Precursor.Normalised)) |> 
  mutate(medium_intensity = ifelse(between(Precursor.Normalised, 
    quantile(Precursor.Normalised,.25), quantile(Precursor.Normalised,.75)) , T, F ) )

medium_intensity_rt <- precursor_int_df |> 
  filter(medium_intensity) |> summarize(RT.Width = round(mean(RT.Width),3))

precursor_RT <- ggplot(precursor_int_df, aes(x = Precursor.Normalised, y = RT.Width*60)) + 
  geom_density2d_filled(alpha = 0.7) +
  geom_hline(yintercept = medium_intensity_rt$RT.Width*60, linetype = "dashed", color = "darkred") +
  annotate("label", x = min(precursor_int_df$Precursor.Normalised), 
                   y = medium_intensity_rt$RT.Width*60, label = paste0("Mean RT: ", medium_intensity_rt$RT.Width*60), vjust = -1, hjust = -2) +
  labs(y = "RT width (sec)", caption =  bquote(bold("Note:")~ "Use peptides wtih Q1-Q3 intensity for mean RT")) +
  scale_color_viridis() +
  theme_classic() +
  theme(plot.caption = element_text(hjust=0))

ggsave(precursor_RT, filename = paste0("result/plot/precursor_RT_", exp_name, ".tiff"), width = 8, height = 5, create.dir = T)
message("Done!")
