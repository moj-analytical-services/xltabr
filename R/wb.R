wb_initialise <- function(tab, wb, ws_name) {

  tab$wb <- wb

  if (!(sheetExists(tab$wb, ws_name))) {
    openxlsx::addWorksheet(tab$wb, ws_name, gridLines = FALSE)
  }

  tab
}
