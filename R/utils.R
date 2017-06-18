sheetExists <- function(wb, ws_name) {
  sheets <- openxlsx::sheets(wb)
  return(ws_name %in% sheets)
}

df_width <- function(tab) {
  ncol(tab$data$df_final)
}
