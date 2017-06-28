#' Create all the required properties for the top headers on the tab object
body_initialise <- function(tab) {

  tab$body <- list()

  # body_df contains the original data which was passed to to body,
  # augmented with two additional columns that store styling information, called
  # meta_row_ and meta_left_header_row_
  tab$body$body_df<- NULL
  tab$body$body_df_to_write <- NULL #This is the df that's actually written to wb

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
  tab$body$body_df$meta_row_ <- "body" # Fill whole col
  tab$body$body_df$meta_left_header_row_ <- "left_header_1" #Fill whole col

  # The meta_ columns deal with styling that varies by row.
  # We also need to set up a vector that contains a style_name per column
  n <- names(tab$body$body_df_to_write)
  tab$body$meta_col_ <- rep("", length(n))
  names(tab$body$meta_col_) <- n
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
    m <- rbind(names(df), col_style_names)
    col_style_names <-m[2,]
    tab$body$meta_col_ <- col_style_names
    tab$body$user_provided_col_style_names <- col_style_names
  }

  tab
}

body_get_wb_cols <- function(tab) {

  if (is.null(tab$body$body_df)) {
    return(integer(0))
  }

  tlc <- tab$extent$topleft_col

  cols_vec <- seq_along(tab$body$body_df_to_write)

  wb_cols <- seq_along(cols_vec) + tlc - 1

  wb_cols


}

body_get_wb_left_header_cols <- function(tab){

  if (is.null(tab$body$body_df)) {
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



body_get_bottom_wb_row <- function(tab) {
  top_headers_bottom <- top_headers_get_bottom_wb_row(tab)
  b_rows <- body_get_wb_rows(tab)

  max(c(top_headers_bottom, b_rows))

}

body_get_rightmost_wb_col <- function(tab) {
  rightmost_th <- top_headers_get_rightmost_wb_col(tab)
  b_cols <- body_get_wb_cols(tab)

  max(c(rightmost_th, b_cols))
}

#' Create table |row|col|style name| containing the styles names
body_get_cell_styles_table <- function(tab) {

  r <- xltabr:::body_get_wb_rows(tab)
  c <- xltabr:::body_get_wb_cols(tab)
  df <- expand.grid(row = r, col = c)

  #All cells get body
  df_br <- data.frame(row = r, body_style = tab$body$body_df$meta_row_)
  df <- merge(df, df_br, by="row")

  #Left headers only
  hc_cols <- xltabr:::body_get_wb_left_header_cols(tab)

  if (length(hc_cols) > 0) {
    df_hc_r <- data.frame(row = r, left_header_style = tab$body$body_df$meta_left_header_row_, stringsAsFactors = FALSE)
    hcs <- data.frame(col = hc_cols)
    df_hcs <- merge(df_hc_r, hcs)

    df <- merge(df, df_hcs, by=c("row", "col"), all.x = TRUE)
    df$left_header_style[is.na(df$left_header_style)] <- ""
  }

  #Columns
  df_th <- data.frame(col = c, top_header_style = tab$body$meta_col_)
  df <- merge(df, df_th, by="col")




  df$style_name <- paste(df$body_style, df$left_header_style, df$top_header_style, sep="|")
  df <- df[, c("row", "col", "style_name")]
  df$style_name <- remove_leading_trailing_pipe(df$style_name)

  df

}

body_write_rows <- function() {
}
