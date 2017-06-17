#' Create a new xltabr object for cross tabulation
#'
#' @param df The dataframe you want to output to Excel
#'
#' @return A list which contains the dataframe
#' @importFrom magrittr %>%
#' @name %>%#'
#' @export
initialise <- function(df, wb = openxlsx::createWorkbook(), ws_name = "Sheet1", title = NULL, startcell_row = 1, startcell_column = 1) {
  tab <- list()
  tab$data <- list()
  tab$metadata <- list()
  tab$wb <- wb

  # Default values - functions are provided for the user to overwrite these
  tab$metadata$num_header_rows <- 1
  tab$metadata$num_header_columns <- 1

  tab$metadata$startcell_row <- startcell_row
  tab$metadata$startcell_column <- startcell_column

  tab$styles <- list()
  tab <- add_default_styles(tab)

  tab$data$title <- title
  tab$data$ws_name <- ws_name
  tab$data$df_orig <- df

  tab <- generate_df_final(tab)

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
