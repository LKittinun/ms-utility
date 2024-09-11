library(QFeatures)

suppressMessages( options(warn=-1) )
args <- commandArgs(trailingOnly = T)

folder_path <- "C:/DIA-NN/1.9/Result/2_Plasma_cervix_ziptip_3k_27_06_2024/200ng/GPF_DDA_search"

exp_name <- tail(unlist(strsplit(folder_path, split = "/")), 1)

setwd(folder_path)
dfq <- readQFeaturesFromDIANN(df)
BiocManager::install("rawrr")
names(dfq)
dfq[[1]]
colnames(rowData(dfq[[1]]))
dfq[[1]] |> 
  rowData() |> 
  as_tibble() |> 
  ggplot(aes(x = RT, y = Mass.Evidence)) + geom_point() +
  geom_hline(yintercept = 10, linetype = "dotdash", col = "red") +
  geom_hline(yintercept = 5, linetype = "dotdash", col = "red") +
  theme_classic()
dfq[[2]] |> 
  rowData() |> 
  as_tibble() |> pull(Mass.Evidence) 
rownames(dfq[["Ms1.Area"]])
library(rawrr)
alb <- as.numeric(readLines("albumin_mass.txt"))
file <- "Z:/Proteomics/C20531700/1_Test_ziptip_3k10k_14_06_2024/Cell/100ng/P16_25_10k_DIA.raw"

test <- readIndex(file)
head(test)
header <- readFileHeader(file) 
spec <- readSpectrum(file, 1:10000)

purrr::map_if(header, \(x) length(x)>1 , \(x)paste(x, collapse = ", ")) |> 
  as.data.frame()

df <- purrr::map(spec , ~data.frame(
  scan = .x[["scanType"]] , TIC = .x[["TIC"]], IIT = .x[["Ion Injection Time (ms):"]],
  RT = .x[["StartTime"]])) |>
  list_rbind() |> 
  mutate(across(c(TIC,IIT,RT), as.numeric))
df |> filter(!grepl("ms2", scan) ) |> 
  ggplot( aes(x = RT, y = IIT)) + geom_point(size = 0.3, alpha = 0.5) + theme_bw()

df |> filter(grepl("ms2", scan) ) |> 
  ggplot( aes(x = RT, y = IIT)) + geom_point(size =0.3, alpha = 0.5) + theme_bw()

plot(spec[[1]], centroid = T, SN = T, diagnostic = T)
abline(h = 5, lty = 2, col = "blue")
yIonSeries <- fragmentIon("LGGNEQVTR")[[1]]$y[1:8]
abline(v = yIonSeries, col = "#DDDDDD88", lwd =5)
axis(3, yIonSeries,paste0("y", seq(1, length(yIonSeries))))
spec[[2]]$scan
scanNumber(spec[[1]])
library(protViz)
library(ExperimentHub)
readChromatogram(file, mass = peptide/2, tol = 5,  type = "xic") |> 
  plot()

PEG <- sapply(1:100, \(x) paste(rep("A", x), collapse = ""))
peptide <- parentIonMass(PEG)
fragmentIon("LGGNEQVTR")[[1]]$y[1:8]

