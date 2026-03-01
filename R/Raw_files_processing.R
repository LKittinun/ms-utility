suppressPackageStartupMessages({
  library(tidyverse)
  library(rawrr)
  library(scales)
  library(furrr)
  })

suppressMessages( options(warn=-1) )

loc <- getwd()
setwd(loc)

#args <- commandArgs(trailingOnly = T)
folder_path <- "Z:/Proteomics/C20530889/MSC"

cont_df <- read_csv("contaminant_skyline.csv",show_col_types = FALSE) |> 
  filter(grepl("peg|triton|polysiloxane|tween", cont , ignore.case=T)) |> 
 # filter(cont %in% c("PEG", "PPG", "Tween", "Polysiloxane", "Triton", "Triton, reduced", "TPO") ) |> 
  mutate(mass = as.character(mass)) 

  cont_df |> 
    filter(grepl("polysiloxane", cont, ignore.case=T))

setwd(folder_path)

file_list_all <- list.files(folder_path, pattern = "*.raw$") |> 
  stringr::str_sort(numeric = T)

dmg_files <- file.size(file_list_all) < 50e6

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

library(furrr)
plan(multisession, workers = 6)

cont_list <- 
  future_map(file_list, .progress = T, \(x)

  { library(rawrr)
    library(tidyverse)
    library(scales)

    xic <- rawrr::readChromatogram(x, unique(cont_df$mass), tol =20, type = "xic")  
    bpc <- rawrr::readChromatogram(x, type = "bpc") 
    tic <- rawrr::readChromatogram(x, type = "tic")
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

    y_ranges<- with(chrom_df, 10^(floor(min(log10(intensities[intensities!=0]))):20))
  
    p <- chrom_df |>
    ggplot( aes(x = times, y = intensities, col = cont)) + 
    geom_line(alpha = 0.5) + 
    geom_line(data = bpc_df, aes(x = times, y = bpc_intensities), 
                                 col = "black", inherit.aes = F, alpha = 0.2) +
    scale_color_manual(values = unname(Polychrome::glasbey.colors(20)[-1])) +
    scale_x_continuous(breaks = seq(0,max(chrom_df$times),2)) +
 #   scale_y_continuous(limits = c(0,1e8), labels=\(x) {ifelse(x!=0,parse(text = gsub("e\\+?", " %*% 10^", scales::scientific_format()(x))), "0")} ) +
    labs(title = x, 
      #caption = "*Grey line represents 5% BPC at given time",
         x= "Time", y = "Intensity", col = "Contaminant") +
    theme_classic() 
  return(list(chrom_df_sum, p, chrom_df))
  }
)  |> 
  set_names(file_list)

## Indiv plot
if (!dir.exists("./raw_summary/cont/ind")) {dir.create("./raw_summary/cont/ind")}
iwalk(cont_list, .progress=T, ~{
  filename <- paste0("raw_summary/cont/ind/", .y, ".png")
  if(!file.exists(filename)){
  ggsave(filename, .x[[2]], dpi = 300, height = 5, width = 10)
  }
})



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
  scale_y_log10(breaks =  trans_breaks("log10", function(x) 10^x), 
                labels = trans_format("log10", math_format(10^.x)),
                guide = "axis_logticks",
                name = "Contaminant Intensity", sec.axis = sec_axis(~., labels = NULL, name = "Total TIC"))+ 
  coord_cartesian(ylim = c(1e6, max(ttic_df$ttic)), expand = T) +
  labs(title = "Intensity plot for each contaminant (log10 scale)",x = "Sample", col = "Contaminant", linetype = "Contaminant") +
  theme_classic() +
  theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust=1))

ggsave("./raw_summary/cont/all_cont_p.png", all_cont_p, height = 4.5, width = max(5,floor(30 / (2 + exp(-0.1 * (length(file_list_all) - 10))))))



if (!dir.exists("./raw_summary/IIT")) {dir.create("./raw_summary/IIT")}

library(furrr)
plan(multisession, workers = 6)


IIT_plots <- future_map(file_list, .progress= T, ~{
  header <- readTrailer(.x)
  read_list <- c("Scan Description:", "Scan Event:", "AGC Target:", "AGC Fill:",
               "Ion Injection Time (ms):" , "Max. Ion Time (ms):" )
  scans <- map(read_list, .y = .x, ~data.frame(.x = readTrailer(.y, .x))) |> 
    list_cbind() |> 
    set_names(read_list)

  index <- readIndex(.x)

  IIT <- cbind(scans,index)
  IIT <- type.convert(IIT, as.is = T)|> 
    janitor::clean_names() |> 
    mutate(ms_order = stringr::str_to_upper(ms_order))
  data.table::fread("D:/Small molecule/Data/16-05-2024_plasma_urine_pos_neg/30-04-2024.sld")
  maxIIT <- IIT |> 
    filter(ms_order == "MS") |> 
    summarize(max = max(ion_injection_time_ms)) 

  p <- ggplot(IIT, aes(x = start_time, y = ion_injection_time_ms, col = ms_order)) +
    geom_point(size = 0.05, aes(alpha = agc_fill)) +
    geom_smooth(col = "black", aes(group = ms_order), se = T, linewidth = 0.1) +
    labs(title = .x, x = "Time (min)", y = "Ion injection time (ms)", col = "Spectrum") +
      scale_alpha_continuous(range = c(0.01,0.1))+
      ggsci::scale_color_lancet()+
      guides(alpha = "none")+
      theme_classic() +
      theme(legend.position="top")

  ggsave(paste0("raw_summary/IIT/plot",.x,".png"), p , create.dir = T)
 
  return(list(IIT,p))
})

cycle_plots <- future_map(file_list, .progress= T, ~{
  header <- readTrailer(.x)
  read_list <- c("Scan Description:", "Scan Event:", Cy,
               "Ion Injection Time (ms):" , "Max. Ion Time (ms):" )
  scans <- map(read_list, .y = .x, ~data.frame(.x = readTrailer(.y, .x))) |> 
    list_cbind() |> 
    set_names(read_list)

  index <- readIndex(.x)

  IIT <- cbind(scans,index)
  IIT <- type.convert(IIT, as.is = T)|> 
    janitor::clean_names() |> 
    mutate(ms_order = stringr::str_to_upper(ms_order))
  
  maxIIT <- IIT |> 
    filter(ms_order == "MS") |> 
    summarize(max = max(ion_injection_time_ms)) 

  p <- ggplot(IIT, aes(x = start_time, y = ion_injection_time_ms)) +
    geom_point(size = 1, shape = 21, aes(alpha = agc_fill, col = ms_order)) +
    geom_smooth(aes(group = ms_order), se = T, linewidth = 0.5) +
    labs(title = .x, x = "Time (min)", y = "Ion injection time (ms)", col = "Spectrum") +
      scale_alpha_continuous(range = c(0.01,0.1))+
      ggsci::scale_color_lancet()+
      guides(alpha = "none")+
      theme_classic() +
      theme(legend.position="top")

  ggsave(paste0("raw_summary/IIT/plot",.x,".png"), p , create.dir = T)
 
  return(list(IIT,p))
})

ggplot(IIT_plots[[1]][[1]], aes(x = start_time, y = ion_injection_time_ms)) +
    geom_point(size = 0.5, shape = 21, alpha = 0.5, aes(alpha = agc_fill, col = ms_order)) +
    geom_smooth(col = "black", aes(group = ms_order), se = T, linewidth = 0.5) +
    labs( x = "Time (min)", y = "Ion injection time (ms)", col = "Spectrum") +
#      scale_shape(solid = F) +
      scale_alpha_continuous(range = c(0,1)) +
      ggsci::scale_color_lancet()+
      guides(alpha = "none")+
      theme_bw() 

ggsave('iit.png', height = 5, width = 7)
tic_df<- readChromatogram(file_list[[1]], type="tic") |> 
 with( data.frame(start_time = as.numeric(times), tic = intensities))
ggplot(tic_df, aes(x = as.numeric(start_time), y = tic)) + geom_line(aes(group = 1))

df<-IIT_plots[[1]][[1]] |> 
  filter(ms_order == "MS") |> 
  mutate(cycle_time = (start_time-lag(start_time))*60) |> 
  fuzzy_left_join(tic_df, by = c("start_time" = "start_time"), 
  match_fun = list(start_time = function(x, y) abs(x-y)<= 0.01)) 


ggplot(df,aes(x = start_time.x, y = cycle_time)) + 
  geom_point(aes(col = tic),size  =1,alpha=0.5, shape = 21)+
  geom_smooth(linewidth = 0.5, se = T) +
  labs(x = "Retention time (min)", y  = "Cycle time (ms)", col = "TIC") +
  scale_y_continuous(breaks = scales::breaks_extended(8))+
  scale_color_viridis_c(labels = \(x){ifelse(x!=0,parse(text = gsub("e\\+?", " %*% 10^", scales::scientific_format()(x))), "0")} ) +
  theme_bw()

readTrailer(file_list[[1]])
test<-rawDiag::readRaw(file_list[[1]])
rawDiag::plotCycleTime(test)
p_trans <- p+
#  transition_states(start_time, state_length = 0)
  transition_manual(start_time,cumulative =  T) 
install.packages("lattice")
library(lattice)
head(volcano)

animation <- animate(p_trans, fps = 100, duration = 60)
anim_save("animated.gif" ,animation = animation)

