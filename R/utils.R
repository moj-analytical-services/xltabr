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

is_null_or_blank <- function(x) {
  if (length(x) > 1) {
    Stop("You passed something of length > 1 to is_null_or_blank")
  }

  if (x == "") {
    return(TRUE)
  }

  if (is_null(x)) {
    return(TRUE)
  }

  return(FALSE)

}
