#' Create all the required properties for the title on the tab object
title_initialise <- function(tab) {

  tab$title <- list()

  tab$title$title_text <- NULL
  tab$title$title_style_names <- NULL
  tab
}

#' Add titles to the tab.  Title text is provided as a character vector, with each element being a row of the title
#'
#' @param title_text A character vector.  Each element is a row of the title
#' @param title_style_names A character vector.  Each elemment is a style_name
#' @export
add_title <- function(tab, title_text, title_style_names) {

  tab$title$title_text <- title_text
  tab$title$title_style_names <- title_style_names

  tab


}

#' Get the column occupied by the title
title_get_wb_cols <- function(tab) {

  num_rows <- length(tab$title$title_text)

  #TODO: If body exists, then use body_get_wb_cols here
  tlc <- tab$extent$topleft_col

  if (num_rows == 0) {
    title_cols = integer(0)
  } else {
    title_cols <- tlc
  }

  title_cols
}

#' Get the rows occupied by the title
title_get_wb_rows <- function(tab) {

  num_rows <- length(tab$title$title_text)

  tlr <- tab$extent$topleft_row
  tlc <- tab$extent$topleft_col

  if (num_rows == 0) {
    title_rows = integer(0)
  } else {
    title_rows = tlr:(tlr + num_rows-1)
  }

  title_rows

}

#' Create table |row|col|style name| containing the styles names
title_get_style_names <- function(tab) {

  rows <- title_get_wb_rows(tab)
  styles <- tab$title$title_style_names
  cols <- title_get_wb_cols(tab)


  df <- data.frame(row = rows, style_name = styles)
  df <- merge(df, data.frame(col = cols))
  df
}


title_write_rows <- function(tab) {
  # For each title, write and apply style

  start_row <- min(title_get_wb_rows(tab))
  start_column <- min(title_get_wb_cols(tab))
  ws_name <- tab$misc$ws_name

  # This will recycle if necessary - e.g. if a single stitle style is provided
  pairs <- cbind(tab$data$title_text, tab$data$title_style_names)

  nrows <- nrow(pairs)
  row_counter <- 0

  if (nrows > 0) {
    for (row_num in 1:nrows) {
      pair <- pairs[row_num,]

      write_row <- start_row + row_counter
      write_col <- start_column

      title <- pair[1]
      style <- tab$style_catalogue[[pair[2]]]$style
      row_height <- tab$style_catalogue[[pair[2]]]$row_height

      openxlsx::writeData(tab$wb, ws_name, title, startRow = write_row, startCol = start_column)

      tab$metadata$rows_before_df_counter <- tab$metadata$rows_before_df_counter + 1
      row_counter = row_counter + 1
    }
  }

  tab

}

