#' Create all the required properties for the top headers on the tab object
body_initialise <- function(tab) {

  tab$body <- list()

  # The top headers are a list of character vectors, each one representing a row
  tab$body$body_df_orig <- NULL
  tab$body$body_df_to_write <- NULL

  tab
}


#' Add a table body (a df) to the tab.
#'
#' @param df For a single top_header row, a character vector.  For multiple top_header rows, a list of character vectors.
#' @param left_header_colnames The columns in the df which are left headers
#'
#' @export
add_body <- function(tab, df, left_header_colnames = NULL, row_style_names = NULL, left_header_style_names = NULL, col_style_names = NULL) {

  tab$body$body_df <- df
  tab$body$body_df_to_write <- df
  tab$body$left_header_colnames <- left_header_colnames

  # Establish 'meta' columns, which contain style names for each row (and will not be written to the wb)
  tab$body$body_df$meta_row_ <- "body" # Fill whole
  tab$body$body_df$meta_left_header_row_ <- "left_header_1"

  # Set up a vector that contains a style_name per column
  tab$body$meta_col_ <- ""

  # If the user has provided row or col style names, add them.  Use recycling rules.
  if (not_null(row_style_names)) {

    m <- rbind(tab$body$body_df[[1]], row_style_names)
    row_style_names <-m[2,]
    tab$body$body_df$meta_row_ <- row_style_names
    tab$body$user_provided_row_style_names <- row_style_names
  }

  if (not_null(left_header_style_names)) {
    m <- rbind(tab$body$body_df[[1]], left_header_style_names)
    left_header_style_names <-m[2,]
    tab$body$body_df$meta_left_header_row_ <- left_header_style_names
    tab$body$user_provided_left_header_style_names <- left_header_style_names
  }

  if (not_null(col_style_names)) {
    m <- rbind(names(df)), col_style_names)
    col_style_names <-m[2,]
    tab$body$meta_col_ <- col_style_names
    tab$body$user_provided_col_style_names <- col_style_names
  }


}

body_get_wb_cols <- function(tab) {
  
  if (is.null(tab$body$body_df_orig)) {
    return(integer(0))
  }

  tlc <- tab$extent$topleft_col

  cols_vec <- seq_along(tab$body$body_df_to_write)

  wb_cols <- seq_along(cols_vec) + tlc - 1

  wb_cols


}

body_get_wb_left_header_cols <- function(tab){ 

  if (is.null(tab$body$body_df_orig)) {
    return(integer(0))
  }

  tlc <- tab$extent$topleft_col

  cols_vec <- seq_along(tab$body$left_header_colnames)

  wb_cols <- seq_along(cols_vec) + tlc - 1

  wb_cols


}

body_get_wb_rows <- function(tab) {
  offset <- top_headers_get_bottom_wb_row(tab)
  seq_along(tab$body$body_df[[1]]) + offset
}



body_get_bottom_wb_row <- function() {
}

body_get_rightmost_wb_col <- function() {
}

body_get_cell_styles_table <- function() {
}

body_write_rows <- function() {
}
