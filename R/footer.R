# Create all the required properties for the footer on the tab object
footer_initialise <- function(tab) {

    tab$footer <- list()
    tab$footer$footer_text <- NULL
    tab$footer$footer_style_names <- NULL

    tab

}

#' Add footers to the tab.  Footer text is provided as a character vector, with each element being a row of the footer
#'
#' @param tab The core tab object
#' @param footer_text A character vector.  Each element is a row of the footer
#' @param footer_style_names A character vector.  Each elemment is a style_name
#' @export
add_footer <- function(tab, footer_text, footer_style_names = "footer") {

   tab$footer$footer_text <- footer_text
   tab$footer$footer_style_names <- footer_style_names

    tab

}

# Get the columns occupied by the footer in the wb
footer_get_wb_cols <- function(tab) {


    #TODO: If body exists, then use body_get_wb_cols here, because this means we can merge title cells to the right width
    num_rows <- length(tab$footer$footer_text)

    tlc <- tab$extent$topleft_col

    if (num_rows == 0) {
        footer_cols = NULL
    } else {
        footer_cols <- tlc
    }

    footer_cols

}


# Get the rows occupied by the footer in the wb
footer_get_wb_rows <- function(tab) {

  offset <- body_get_bottom_wb_row(tab)

  num_rows <- length(tab$footer$footer_text)


  if (num_rows == 0) {
    footer_rows = NULL
  } else {
    footer_rows = seq_along(tab$footer$footer_text) + offset
  }

  footer_rows
}

# Get the bottom row of the footer in the wb
footer_get_bottom_wb_row <- function(tab) {

    body_bottom <- body_get_bottom_wb_row(tab)
    footer_rows <- footer_get_wb_rows(tab)

    max(c(body_bottom, footer_rows))
}

# Get the rightmost column of the footers in the wb
footer_get_rightmost_wb_col <- function(tab) {
    footer_cols <- footer_get_wb_cols(tab)
    max(c(footer_cols, tab$extent$topleft_col - 1))
}

# Create table |row|col|style name| containing the styles names
footer_get_cell_styles_table <- function(tab) {

  rows <- footer_get_wb_rows(tab)

  if (length(rows) == 0) {
    df <- data.frame("row" = integer(0), "col" = integer(0), "style_name" = character(0), stringsAsFactors = FALSE)
    return(df)
  }

  rows <- footer_get_wb_rows(tab)
  styles <- tab$footer$footer_style_names

  if (not_null(tab$body)) {
    cols <- body_get_wb_cols(tab)
  } else {
    cols <- footer_get_wb_cols(tab)
  }

  df <- data.frame(row = rows, style_name = styles, stringsAsFactors = FALSE)
  df <- merge(df, data.frame(col = cols, stringsAsFactors = FALSE)) #merge with no join column creates cartesian product
  df <- df[,c("row", "col", "style_name")]
  df
}

# Write all the required data (but no styles)
footer_write_rows <- function(tab) {

  if (is.null(tab$footer$footer_text)) {
    return(tab)
  }

  ws_name <- tab$misc$ws_name

  col <- min(footer_get_wb_cols(tab))
  rows <- footer_get_wb_rows(tab)
  counter <- 1
  for (r in rows) {
    footer <- tab$footer$footer_text[counter]
    counter = counter + 1
    openxlsx::writeData(tab$wb, ws_name, footer, startRow = r, startCol = col)
  }

  tab

}
