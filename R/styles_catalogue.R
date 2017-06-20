# Iterate through the default style workbook, adding them to the styles catalogue
add_default_styles <- function (tab) {
  path <- system.file("extdata", "styles.xlsx", package = "xltabr")
  wb <- openxlsx::loadWorkbook(path)
  openxlsx::readWorkbook(wb)

  for (i in wb$styleObjects) {
    r <- i$rows
    c <- i$cols
    cell <- openxlsx::readWorkbook(wb, rows = r, cols = c, colNames = FALSE, rowNames = FALSE)
    value <- cell[1, 1]

    tab$style_catalogue[[value]] <- list()
    tab$style_catalogue[[value]]$style <- i$style

    cell <- openxlsx::readWorkbook(wb, rows = r, cols = c + 1, colNames = FALSE, rowNames = FALSE)
    tab$style_catalogue[[value]]$row_height <- cell[1, 1]
  }

  tab
}


#' @export
override_styles <- function(tab) {
  # Code here to override styles from Excel document provided by user
}

