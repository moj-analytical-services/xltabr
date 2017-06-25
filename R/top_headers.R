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
#' @param top_header For a single top_header row, a character vector.  For multiple top_header rows, a list of character vectors.
#' @param row_style_names A character vector, with an element for each row of the top header.  Each element is a style_name (i.e. a key in the style catalogue)
#' @param col_style_names A character vector, with and elemment for each column of the top header.  Each element is a style name. Col styles in inherit from row_styles.
#'
#' @export
add_top_headers <- function(tab, top_headers, col_style_names=NULL, row_style_names="body|top_header_1") {

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

}
