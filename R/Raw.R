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

# {
# proton <- 1.007276
# PEG <- c(63.0441, 107.0703, 151.0965, 195.1227, 239.149, 283.1752, 327.2014, 371.2276, 415.2538, 459.28, 503.3062, 547.3325, 591.3587, 635.3849, 679.4111, 723.4373, 767.4635, 811.4897, 855.516, 899.5422)
# TritonX100 <- c(251.2006, 295.2268, 339.253, 383.2792, 427.3054, 471.3316, 515.3578, 559.384, 603.4102, 647.4364, 691.4626, 735.4888, 779.515, 823.5412, 867.5674, 911.5936, 955.6198, 999.646, 1043.672, 1087.698)
# TritonX100_red <- c(257.2476, 301.2738, 345.3, 389.3262, 433.3524, 477.3786, 521.4048, 565.431, 609.4572, 653.4834, 697.5096, 741.5358, 785.562, 829.5882, 873.6144, 917.6406, 961.6668, 1005.693, 1049.719, 1093.745)
# TritonX101 <- c(265.2163, 309.2425, 353.2687, 397.2949, 441.3211, 485.3473, 529.3735, 573.3997, 617.4259, 661.4521, 705.4783, 749.5045, 793.5307, 837.5569, 881.5831, 925.6093, 969.6355, 1013.662, 1057.688, 1101.714)
# TritonX101_red <- c(271.2632, 315.2894, 359.3156, 403.3418, 447.368, 491.3942, 535.4204, 579.4466, 623.4728, 667.499, 711.5252, 755.5514, 799.5776, 843.6038, 887.63, 931.6562, 975.6824, 1019.709, 1063.735, 1107.761)
# Polysiloxane <- c(75.026, 149.0448, 223.0636, 297.0824, 371.1012, 445.12, 519.1388, 593.1576, 667.1764, 741.1952, 815.214, 889.2328, 963.2516, 1037.27, 1111.289, 1185.308, 1259.327, 1333.345, 1407.364, 1481.383)

# mass_list <- map(list("PEG","TritonX100","TritonX100_red","TritonX101","TritonX101_red","Polysiloxane"), 
# ~data.frame(cont = .x, h1 = get(.x)) |> 
#   mutate(h2 = (h1 + proton)/2) |> 
#   pivot_longer(c(h1, h2), names_to = "charge", values_to = "mass") |> 
#   mutate(mass = as.character(mass))) |> 
#   list_rbind()
# }