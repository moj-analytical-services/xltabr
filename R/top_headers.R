#' Create all the required properties for the top headers on the tab object
top_headers_initialise <- function(tab) {

  tab$top_headers <- list()

  # The top headers are a list of character vectors, each one representing a row
  tab$top_headers$top_headers_list <- NULL
  tab$top_headers$top_headers_row_style_names <- NULL
  tab$top_headers$top_headers_col_style_names <- NULL
  tab
}

#' Add top headers to the tab.  The top headers are provided as a character vector.
#' If you need more than one row, provide a list of character vectors.
#' Top headers are automatically assigned the style_text 'top_header_1', but
#' you may provide style overrides using column_style_names and row_style_names
#'
#' @param tab The core tab object
#' @param top_header For a single top_header row, a character vector.  For multiple top_header rows, a list of character vectors.
#' @param row_style_names A character vector, with an element for each row of the top header.  Each element is a style_name (i.e. a key in the style catalogue)
#' @param col_style_names A character vector, with and elemment for each column of the top header.  Each element is a style name. Col styles in inherit from row_styles.
#'
#' @export
add_top_headers <- function(tab, top_headers, col_style_names="", row_style_names="body|top_header_1") {

  # Check types and assign the data to top_headers_list.  Each row is an element in the list
  if (typeof(top_headers) == "character") {
    top_headers <- list(top_headers)
  }
  tab$top_headers$top_headers_list <- top_headers


  # Use recycling to make the col_style_names and row_style_names match the number of entries in top headers
  m <-rbind(top_headers[[1]],col_style_names)
  col_style_names <- m[2,]

  m <-rbind(seq_along(top_headers), row_style_names)
  row_style_names <- m[2,]

  tab$top_headers$top_headers_row_style_names <- row_style_names
  tab$top_headers$top_headers_col_style_names <- col_style_names

  tab

}

top_headers_get_wb_cols <- function(tab) {

  if (is.null(tab$top_headers$top_headers_list)) {
    return(NULL)
  }

  tlc <- tab$extent$topleft_col

  header_cols_vec <- tab$top_headers$top_headers_list[[1]]

  wb_cols <- seq_along(header_cols_vec) + tlc - 1

  wb_cols


}

top_headers_get_wb_rows <- function(tab) {

  if (is.null(tab$top_headers$top_headers_list)) {
    return(NULL)
  }

  offset <- title_get_bottom_wb_row(tab)
  seq_along(tab$top_headers$top_headers_list) + offset


}

top_headers_get_bottom_wb_row <- function(tab) {

  title_bottom <- title_get_bottom_wb_row(tab)
  th_rows <- top_headers_get_wb_rows(tab)

  max(c(title_bottom, th_rows))

}

top_headers_get_rightmost_wb_col <- function(tab) {

  th_cols <- top_headers_get_wb_cols(tab)

  if (length(th_cols) == 0) {
    return(tab$extent$topleft_col - 1)
  } else {
    return(max(th_cols))
  }


}

#' Create table |row|col|style name| containing the styles names
top_headers_get_cell_styles_table <- function(tab) {

  rows <- top_headers_get_wb_rows(tab)

  if (length(rows) == 0) {
    df <- data.frame("row" = integer(0), "col" = integer(0), "style_name" = character(0), stringsAsFactors = FALSE)
    return(df)
  }

  rs <- tab$top_headers$top_headers_row_style_names
  df1 <- data.frame(rs, row = rows, stringsAsFactors = FALSE)

  cs <- tab$top_headers$top_headers_col_style_names
  cols <-  top_headers_get_wb_cols(tab)

  df2 <- data.frame(cs, col = cols, stringsAsFactors = FALSE)
  df <- merge(df1, df2)

  df$style_name <- paste(df$rs, df$cs, sep="|")
  df$style_name <- remove_leading_trailing_pipe(df$style_name)

  df[, c("row", "col", "style_name")]


}

#' Write all the required data (but no styles)
top_headers_write_rows <- function(tab) {

  if (is.null(tab$top_headers$top_headers_list)) {
    return(tab)
  }

  ws_name <- tab$misc$ws_name

  #TODO:check there's something to write before writing

  #The transpose operation is safe because we want everything to be character
  data <- t(data.frame(tab$top_headers$top_headers_list))

  col <- min(top_headers_get_wb_cols(tab))
  row <- min(top_headers_get_wb_rows(tab))


  openxlsx::writeData(tab$wb, ws_name, data, startRow = row, startCol = col, colNames = FALSE)

  tab

}
