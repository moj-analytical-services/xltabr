sheetExists <- function(wb, ws_name) {
  sheets <- openxlsx::sheets(wb)
  return(ws_name %in% sheets)
}

not_null <- function(x) {
  !(is.null(x))
}


remove_leading_trailing_pipe <- function(x) {
  x <-  gsub("\\|+$", "", x, perl=TRUE)
  x <-  gsub("^\\|+", "", x, perl=TRUE)
  x
}
