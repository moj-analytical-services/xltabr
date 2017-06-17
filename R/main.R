#' Create a new xltabr object for cross tabulation
#'
#' @param df The dataframe you want to output to Excel
#'
#' @return A list which contains the dataframe
#' @importFrom magrittr %>%
#' @name %>%#'
#' @export
initialise <- function(df, wb = openxlsx::createWorkbook(), ws_name = "Sheet1", startcell_row = 1, startcell_column = 1) {
  tab <- list()
  tab$data <- list()
  tab$metadata <- list()
  tab$wb <- wb

  # Default values - functions are provided for the user to overwrite these
  tab$metadata$num_header_rows <- 1
  tab$metadata$num_header_columns <- 1

  tab$metadata$startcell_row <- startcell_row
  tab$metadata$startcell_column <- startcell_column

  # Together the following counters keep track of how many rows of space this xltabr takes up
  # A counter which stores how many rows we've written before writing the core df
  tab$metadata$rows_before_df_counter <- 0

  # A counter which stores how many rows we've written before after the core df
  tab$metadata$cols_before_df_counter <- 0


  tab$style_catalogue <- list()
  tab <- add_default_styles(tab)

  tab$styles <- list()
  tab$styles$title_styles <- NULL

  tab$data$title <- title
  tab$data$ws_name <- ws_name
  tab$data$df_orig <- df

  tab <- generate_df_final(tab)

  tab
}

#' @export
add_titles <- function(tab, title = "My title", title_style = "row_header_1") {
  tab$data$title <- title
  tab$data$title_style <- title_style
  tab
}

#' Convert multiple column headers in an xltabr object into
#'
#' @param df The dataframe you want to output to Excel
#'
#' @return A list which contains the dataframe
#' @export
combine_column_headers <- function() {



}
