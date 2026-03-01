
suppressPackageStartupMessages({
  library(tidyverse)
  })
install.packages("XML")
XML::xmlParse("D:/Method/Lin0945p_86min_DDA_120k_30k.meth")

meth_name <- list.files("D:/Method", pattern = "*.meth", full.names = T)

meth_list <- map(meth_list,  .progress = T,
  ~{
    ss <- system(glue::glue('py -m XCaliburMethodReader {.x} -s "TNG-Merkur"'), intern = T) |> 
      stringr::str_trim(side = "both") |> 
      stringi::stri_remove_empty()

    ssexp <- grep("^Experiment \\d", ss)
    sss <- split(ss, findInterval(seq_along(ss), ssexp)) |> 
      set_names(c("global", ss[ssexp]))
    
    exp <- map(sss, .y=.x, ~{
      val <- grep( "=", .x)
      ssval <- .x[val]
      exp_df <- map(strsplit(ssval, split = "="), 
          \(text) {
            text <- str_trim(text)
            setNames(text[2], text[1]) |>
              t() |> 
              as.data.frame() |> 
              janitor::clean_names() }) |> 
        list_cbind(name_repair = "minimal") 
      colnames(exp_df) <- make.unique(colnames(exp_df))
      exp_df <- exp_df |> 
        mutate(method = .y) |> 
        relocate(method , .before = everything())
    }) |> 
      set_names(c("global", ss[ssexp]))
    return(set_names(list(ss, exp), c("raw", "df")) )
  }
) |> 
  set_names(meth_name)

global_meth <- map(meth_list, ~.x[[1]]) |> 
                    list_rbind()
exp_meth <-  map(meth_list, ~.x[c(2:length(.x))] |> 
                  list_rbind()) |>
              list_rbind() 
cat(meth_list[[1]][[1]], sep = "\n")
cat(names(meth_list[1]), sep = "\n")
sink("raw_meth.txt")
iwalk(meth_list, ~{
  cat(.y, sep = "\n")
  cat(.x[[1]], sep = "\n")
  cat("----", sep = "\n")
})
sink(NULL)
write.csv(global_meth, "global_meth.csv", na = "", row.names = F)

write.csv(exp_meth, "exp_meth.csv",  na = "",row.names = F)

