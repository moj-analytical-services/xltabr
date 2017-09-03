#' Create all the required properties for the title on the tab object
title_initialise <- function(tab) {

  tab$title <- list()

  tab$title$title_text <- NULL
  tab$title$title_style_names <- NULL
  tab
}


#' Add titles to the tab.  Title text is provided as a character vector, with each element being a row of the title
#'
#' @param tab The core tab object
#' @param title_text A character vector.  Each element is a row of the title
#' @param title_style_names A character vector.  Each elemment is a style_name
#' @export
add_title <- function(tab, title_text, title_style_names = NULL) {

  if (is.null(title_style_names)) {
    title_style_names = rep("subtitle", length(title_text))
    title_style_names[1] = "title"
  }

  tab$title$title_text <- title_text
  tab$title$title_style_names <- title_style_names

  tab


}

#' Get the column occupied by the title in the wb
title_get_wb_cols <- function(tab) {

  num_rows <- length(tab$title$title_text)

  #TODO: If body exists, then use body_get_wb_cols here, because this means we can merge title cells to the right width
  tlc <- tab$extent$topleft_col

  if (num_rows == 0) {
    title_cols = NULL
  } else {
    title_cols <- tlc
  }

  title_cols
}

#' Get the rows occupied by the title in the wb
title_get_wb_rows <- function(tab) {

  num_rows <- length(tab$title$title_text)

  tlr <- tab$extent$topleft_row

  if (num_rows == 0) {
    title_rows = NULL
  } else {
    title_rows = tlr:(tlr + num_rows-1)
  }

  title_rows

}

#' Get the bottom row of the titles in the wb
title_get_bottom_wb_row <- function(tab) {
  title_rows <- title_get_wb_rows(tab)

  #-1 because if title doesn't exist, the next item will be instructed to write to the row after the title
  max(c(title_rows, tab$extent$topleft_row - 1))
}

#' Get the rightmost column of the titles in the wb
title_get_rightmost_wb_col <- function(tab) {
  title_cols <- title_get_wb_cols(tab)
  max(c(title_cols, tab$extent$topleft_col -1))
}

#' Create table |row|col|style name| containing the styles names
title_get_cell_styles_table <- function(tab) {

  rows <- title_get_wb_rows(tab)

  if (length(rows) == 0) {
    df <- data.frame("row" = integer(0), "col" = integer(0), "style_name" = integer(0))
    return(df)
  }

  rows <- title_get_wb_rows(tab)

  styles <- tab$title$title_style_names

  if (not_null(tab$body)) {
    cols <- body_get_wb_cols(tab)
  } else {
    xols <- title_get_wb_cols(tab)
  }

  df <- data.frame(row = rows, style_name = styles, stringsAsFactors = FALSE)

  df <- merge(df, data.frame(col = cols)) #merge with no join column creates cartesian product
  df <- df[,c("row", "col", "style_name")]
  df
}

#' Write all the required data (but no styles)
title_write_rows <- function(tab) {

  if (is.null(tab$title$title_text)) {
    return(tab)
  }

  ws_name <- tab$misc$ws_name

  col <- min(title_get_wb_cols(tab))
  rows <- title_get_wb_rows(tab)

  counter <- 1
  for (r in rows) {
    title <- tab$title$title_text[counter]
    counter = counter + 1
    openxlsx::writeData(tab$wb, ws_name, title, startRow = r, startCol = col)
  }

  tab

}

