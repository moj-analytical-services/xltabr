#' Create a new xltabr object for cross tabulation
#'
#' @param df The dataframe you want to output to Excel
#'
#' @return A list which contains the dataframe
#' @importFrom magrittr %>%
#' @name %>%#'
#' @export
initialise <- function(wb = openxlsx::createWorkbook(), ws_name = "Sheet1", styles_xlsx = NULL, num_styles_csv = NULL, topleft_row = 1, topleft_col = 1) {

  # Main object
  tab <- list()

  tab <- extent_initialise(tab, topleft_row, topleft_col)
  tab <- title_initialise(tab)
  tab <- body_initialise(tab)

  tab <- style_catalogue_initialise(tab, styles_xlsx = styles_xlsx, num_styles_csv = num_styles_csv)
  tab <- wb_initialise(tab, wb, ws_name)


  tab$misc$ws_name <- ws_name

  tab


}

initialise_debug <- function(wb = openxlsx::createWorkbook(), ws_name = "Sheet1", topleft_row = 1, topleft_col = 1) {

  # Main object
  tab <- list()

  tab <- extent_initialise(tab, topleft_row, topleft_col)
  tab <- title_initialise(tab)
  tab <- body_initialise(tab)

  tab
}
