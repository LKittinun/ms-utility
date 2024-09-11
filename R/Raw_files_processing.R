suppressPackageStartupMessages({
  library(tidyverse)
  library(rawrr)
  library(scales)
  })

suppressMessages( options(warn=-1) )

loc <- getwd()
setwd(loc)

#args <- commandArgs(trailingOnly = T)
folder_path <- "Z:/Metabolomics/Result/16-05-2024_noexclude"

cont_df <- read_csv("contaminant_skyline.csv",show_col_types = FALSE) |> 
  filter(grepl("peg|triton|polysiloxane|tween", cont , ignore.case=T)) |> 
 # filter(cont %in% c("PEG", "PPG", "Tween", "Polysiloxane", "Triton", "Triton, reduced", "TPO") ) |> 
  mutate(mass = as.character(mass)) 

setwd(folder_path)

file_list_all <- list.files(folder_path, pattern = "*.raw$") |> 
  stringr::str_sort(numeric = T)

dmg_files <- file.size(file_list_all) < 10e6

file_list <- file_list_all[!dmg_files]

{
if(sum(dmg_files !=0)) {
  message(paste(file_list[dmg_files], collapse = ", "), " skipped due to too low file size, probably damaged.")
}

metadata <- map(file_list[!dmg_files], .progress = T, {
  possibly(\(i) {
    header <- rawrr::readFileHeader(i)  
    meta <- map_if(header,  \(x) length(x)>1 , \(x) paste(x, collapse = ", ") )
    as.data.frame(meta)
  }, otherwise = NA_real_)
}
) 
meta_df <- list_rbind(keep(metadata, ~is.data.frame(.x)))

if (!dir.exists("./raw_summary")) {dir.create("./raw_summary")}

write_delim(meta_df, file= "./raw_summary/sample_meta.tsv", delim = "\t")
}


if (!dir.exists("./raw_summary/cont")) {dir.create("./raw_summary/cont")}

cont_list <- 
  map(file_list, .progress = T, \(x)
  { xic <- readChromatogram(x, unique(cont_df$mass), tol =20, type = "xic")  
    bpc <- readChromatogram(x, type = "bpc") 
    tic <- readChromatogram(x, type = "tic")
    bpc_df <- with(bpc, data.frame(times = as.numeric(times), bpc_intensities = intensities))
    tbpc <- sum(bpc$intensities)
    tic_df <- with(tic, data.frame(times = as.numeric(times), tic_intensities = intensities))
    ttic <- sum(tic$intensities)


  chrom_df <- map(xic, ~with(.x, data.frame(filter = filter, ppm = ppm, mass = as.character(mass), 
    times = times, intensities = intensities))) |> 
    list_rbind() |> 
    left_join(cont_df, by = "mass") 

  chrom_df_sum <- summarise(chrom_df, intensities_total = sum(intensities), ttic = ttic, bpc = tbpc, mass = max(mass),
                                      intensities_perc = round(intensities_total*100/tbpc,3), .by = cont )
  
  p <- chrom_df |>
    ggplot( aes(x = times, y = intensities, col = cont)) + 
    geom_line() + 
    geom_line(data = bpc_df, aes(x = times, y = bpc_intensities), 
                                 col = "black", inherit.aes = F, alpha = 0.2) +
    scale_color_manual(values = unname(Polychrome::glasbey.colors(20)[-1])) +
    scale_x_continuous(breaks = seq(0,max(chrom_df$times),10)) +
    labs(title = x, caption = "*Grey line represents BPC at given time",
         x= "Time", y = "Intensity", col = "Contaminant") +
    theme_classic() 
  
  return(list(chrom_df_sum, p, chrom_df))
  }
)  |> 
  set_names(file_list)

cont_intensity_df <- imap(cont_list, ~.x[[1]] |> 
  mutate(filename = .y) |> 
  mutate(filename_trunc = stringr::str_trunc(filename, 20, side = "left")) |> 
  relocate(filename, .before = everything())) |> 
  na.omit() |> 
  list_rbind() 

write_delim(cont_intensity_df, file= "./raw_summary/cont/1_sample_cont.tsv", delim = "\t")

ttic_df <- cont_intensity_df |> 
  distinct(filename_trunc, ttic)

all_cont_p <- cont_intensity_df |>
  ggplot(aes(x = filename_trunc)) + 
  geom_line(aes( group = cont, y = intensities_total , col = cont, linetype = cont), linewidth = 0.25) +
  geom_line(data = ttic_df, aes(group = 1, y = ttic), alpha = 0.5, linewidth = 0.9)+
  scale_color_manual(values = unname(Polychrome::glasbey.colors(20)[-1])) +
  scale_y_log10(breaks = 10^seq(0, 15, by = 1), name = "Contaminant Intensity", sec.axis = sec_axis(~., labels = NULL, name = "Total TIC"))+ 
  coord_cartesian(ylim = c(1e6,max(ttic_df$ttic)), expand = T) +
  labs(title = "Intensity plot for each contaminant (log10 scale)",x = "Sample", col = "Contaminant", linetype = "Contaminant") +
  theme_classic() +
  theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust=1))

if (!dir.exists("./raw_summary/cont/ind")) {dir.create("./raw_summary/cont/ind")}

ggsave("./raw_summary/cont/all_cont_p.png", all_cont_p, height = 4.5, width = 20)

iwalk(cont_list, .progress=T, ~{
  filename <- paste0("raw_summary/cont/ind/", .y, ".png")
  if(!file.exists(filename)){
  ggsave(filename, .x[[2]], dpi = 300, height = 5, width = 7)
  }
})

