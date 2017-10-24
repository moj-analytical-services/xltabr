#' Create all the required properties for the title on the tab object
#'
#' @param tab core tab object
title_initialise <- function(tab) {

  tab$title <- list()

  tab$title$title_text <- NULL  # A character vector containing the text of the title, one element per row
  tab$title$title_style_names <- NULL # A character vector containing the style names for each row of the titles
  tab
}


#' Add titles to the tab.  Title text is provided as a character vector, with each element being a row of the title
#'
#' @param tab The core tab object
#' @param title_text A character vector.  Each element is a row of the title
#' @param title_style_names A character vector.  Each elemment is a style_name
#' @export
#' @examples
#' crosstab <- read.csv(system.file("extdata", "example_crosstab.csv", package="xltabr"))
#' tab <- initialise()
#' title_text <- c("Main title on first row", "subtitle on second row")
#' title_style_names <- c("title", "subtitle")
#' tab <- add_title(tab, title_text, title_style_names)
add_title <- function(tab, title_text, title_style_names = NULL) {

  if (is.null(title_style_names)) {
    title_style_names = rep("subtitle", length(title_text))
    title_style_names[1] = "title"
  }

  tab$title$title_text <- title_text
  tab$title$title_style_names <- title_style_names

  tab


}

#' Get the columns occupied by the title in the workbook
#'
#' @param tab The core tab object
title_get_wb_cols <- function(tab) {

  num_rows <- length(tab$title$title_text)

  tlc <- tab$extent$topleft_col

  if (num_rows == 0) {
    title_cols = NULL
  } else {
    title_cols <- tlc
  }

  title_cols
}

#' Get the rows occupied by the title in the workbook
#'
#' @param tab The core tab object
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

#' Get the bottom row which the titles occupy in the workbook
#' If no titles have been provided, return the cell above the topleft row of the extent.
#' The next element (top headers, body or whatever it is) will want to be positioned
#' in the cell below this
#'
#' @param tab The core tab object
title_get_bottom_wb_row <- function(tab) {
  title_rows <- title_get_wb_rows(tab)

  #-1 because if title doesn't exist, the next item will be instructed to write to the row after the title
  max(c(title_rows, tab$extent$topleft_row - 1))
}

#' Get the rightmost column occupied by the titles in the workbook
#'
#' @param tab The core tab object
title_get_rightmost_wb_col <- function(tab) {
  title_cols <- title_get_wb_cols(tab)
  max(c(title_cols, tab$extent$topleft_col -1))
}

#' Create table with columns |row|col|style name| that contains the styles names
#'
#' @param tab The core tab object
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
    cols <- title_get_wb_cols(tab)
  }

  df <- data.frame(row = rows, style_name = styles, stringsAsFactors = FALSE)

  df <- merge(df, data.frame(col = cols)) #merge with no join column creates cartesian product
  df <- df[,c("row", "col", "style_name")]
  df
}

#' Write all the title data to the workbook (but do not write style data)
#'
#' @param tab The core tab object
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

