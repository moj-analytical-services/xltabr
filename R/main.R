#' Create a new xltabr object for cross tabulation
#'
#' @param df The dataframe you want to output to Excel
#' @param insert_below_tab If given, the new tab will be inserted immediately below this one
#'
#' @return A list which contains the dataframe
#' @importFrom magrittr %>%
#' @name %>%#'
#' @export
initialise <- function(wb = NULL, ws_name = NULL, styles_xlsx = NULL, num_styles_csv = NULL, topleft_row = 1, topleft_col = 1, insert_below_tab = NULL) {

  # Main object
  tab <- list()

  if (is.null(ws_name)) {
    ws_name <- "Sheet1"
  }

  if (is.null(wb)) {
    wb <- openxlsx::createWorkbook()
  }


  # If a 'insert below tab' is provided, the user is saying 'put this new table below this existing table'
  if (not_null(insert_below_tab)) {
    ws_name <- insert_below_tab$misc$ws_name
    topleft_row <- extent_get_bottom_wb_row(insert_below_tab) + 1
    topleft_col <- extent_get_rightmost_wb_col(tab)

    # Make sure we don't try to write to an impossible column
    if (topleft_col == 0){
      topleft_col = 1
    }

  }



  tab <- extent_initialise(tab, topleft_row, topleft_col)


  tab <- style_catalogue_initialise(tab, styles_xlsx = styles_xlsx, num_styles_csv = num_styles_csv)

  if (not_null(insert_below_tab)) {
    tab <- wb_initialise(tab, insert_below_tab$wb, ws_name)
  } else {
    tab <- wb_initialise(tab, wb, ws_name)
  }

  tab$misc$ws_name <- ws_name

  tab


}
