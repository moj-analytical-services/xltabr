# First write raw data
#' @export
write_to_wb <- function(tab) {

  ws_name <- tab$data$ws_name

  if (!(sheetExists(tab$wb, ws_name))) {
    openxlsx::addWorksheet(tab$wb, ws_name, gridLines = FALSE)
  }

  tab <- write_title_rows(tab)

  openxlsx::writeData(
    tab$wb,
    tab$data$ws_name,
    tab$data$df_final,
    startRow = tab$metadata$startcell_row + tab$metadata$rows_before_df_counter,
    startCol = tab$metadata$startcell_column
  )


  # tab <- apply_styles_to_wb(tab)

  tab
}


write_title_rows <- function(tab) {
  # For each title, write and apply style

  start_row <- tab$metadata$startcell_row + tab$metadata$rows_before_df_counter
  start_column <- tab$metadata$startcell_column
  ws_name <- tab$data$ws_name

  # This will recycle if necessary
  pairs <- cbind(tab$data$title, tab$data$title_style)

  nrows <- nrow(pairs)
  row_counter <- 0
  write_start_row <-
  if (nrows > 0) {
    for (row_num in 1:nrows) {
      pair <- pairs[row_num,]

      write_row <- start_row + row_counter
      write_col <- start_column

      title <- pair[1]
      style <- tab$style_catalogue[[pair[2]]]$style
      row_height <- tab$style_catalogue[[pair[2]]]$row_height

      openxlsx::writeData(tab$wb, ws_name, title, startRow = write_row, startCol = write_col)
      openxlsx::addStyle(tab$wb, ws_name, style, rows = write_row, cols=write_col)
      openxlsx::setRowHeights(tab$wb, ws_name, rows=write_row, heights=row_height)

      tab$metadata$rows_before_df_counter <- tab$metadata$rows_before_df_counter + 1
      row_counter = row_counter + 1
    }
  }

  tab

}



