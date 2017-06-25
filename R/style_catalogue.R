# Use default .xlsx style catalogue to initialise the catalogue of styles
style_catalogue_initialise <- function (tab) {
  path <- system.file("extdata", "styles.xlsx", package = "xltabr")
  tab <- style_catalogue_xlsx_import(tab, path)
  tab
}

# Iterate through the default style workbook, adding them to the styles catalogue
style_catalogue_xlsx_import <- function(tab, path) {

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

#' Allows the user to provide a .xlsx with custom defined styles
#' Use [this](https://github.com/moj-analytical-services/xltabr/blob/dev/inst/extdata/styles.xlsx?raw=true) file as a basis
#'
#' @param path a file path to the .xlsx file you want to use to override styles
#'
#' @return The tab
#' @export
style_catalogue_override_styles <- function(tab, path_to_xlsx) {
  # Code here to override styles from Excel document provided by user
  tab <- style_catalogue_xlsx_import(tab, path_to_xlsx)
  tab
}

