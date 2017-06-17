# First write raw data
#' @export
write_to_wb <- function(tab) {

  ws_name <- tab$data$ws_name

  if (!(sheetExists(tab$wb, ws_name))) {
    openxlsx::addWorksheet(tab$wb, ws_name)
  }

  openxlsx::writeData(
    tab$wb,
    tab$data$ws_name,
    tab$data$df_final,
    startRow = tab$metadata$startcell_row,
    startCol = tab$metadata$startcell_column
  )

  tab
}
