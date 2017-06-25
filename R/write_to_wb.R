# First write raw data
#' @export
write_to_wb <- function(tab) {

  ws_name <- tab$data$ws_name

  tab <- write_title_rows(tab)

  openxlsx::writeData(
    tab$wb,
    tab$data$ws_name,
    tab$data$df_final,
    startRow = tab$metadata$startcell_row + tab$metadata$rows_before_df_counter,
    startCol = tab$metadata$startcell_column
  )


  apply_styles_to_wb(tab)

  tab
}




