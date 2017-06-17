add_default_styles <- function (tab) {
  paths <- system.file("extdata", "styles.xlsx", package = "xltabr")
  wb <- openxlsx::loadWorkbook(path)
  openxlsx::readWorkbook(wb)

  for (i in wb$styleObjects) {
    r <- i$rows
    c <- i$cols
    cell <- openxlsx::readWorkbook(wb, rows = r, cols = c, colNames = FALSE, rowNames = FALSE)
    value <- cell[1, 1]

    tab$styles[[value]] <- i
  }

  tab
}


#' @export
override_styles <- function(tab) {
  # Code here to override styles from Excel document provided by user
}

