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

#' Get the column occupied by the title in the wb
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

#' Get the rows occupied by the title in the wb
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
title_get_cell_styles_table <- function(tab) {

  rows <- title_get_wb_rows(tab)
  styles <- tab$title$title_style_names
  cols <- title_get_wb_cols(tab)


  df <- data.frame(row = rows, style_name = styles)
  df <- merge(df, data.frame(col = cols))
  df <- df[,c("row", "col", "style_name")]
  df
}

#' Write all the required data (but no styles)
title_write_rows <- function(tab) {

  ws_name <- tab$misc$ws_name

  col <- min(title_get_wb_cols(tab))
  rows <- title_get_wb_rows

  for (r in rows) {
    openxlsx::writeData(tab$wb, ws_name, title, startRow = write_row, startCol = start_column)
  }

  tab
}

