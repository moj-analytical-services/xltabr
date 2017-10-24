#' Create a new xltabr object for cross tabulation
#'
#' @param wb An openxlsx workbook object to write to.  If null, a new one will be created
#' @param ws_name The worksheet to write to.  If null, Sheet1 will be used
#' @param topleft_row Specifies the row where the table begins in the worksheet
#' @param topleft_col Specifies the column where the table begins in the worksheet
#' @param insert_below_tab If given, the new tab will be inserted immediately below this one
#'
#' @return A list which contains the dataframe
#' @importFrom magrittr %>%
#' @export
#' @examples
#' tab <- initialise()
initialise <- function(wb = NULL, ws_name = NULL, topleft_row = 1, topleft_col = 1, insert_below_tab = NULL) {

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
    topleft_col <- extent_get_cols(insert_below_tab)[1]

    # Make sure we don't try to write to an impossible column
    if (is.null(topleft_col)) {
      topleft_col = 1
    }

  }

  tab <- extent_initialise(tab, topleft_row, topleft_col)

  tab <- style_catalogue_initialise(tab)

  if (not_null(insert_below_tab)) {
    tab <- wb_initialise(tab, insert_below_tab$wb, ws_name)
  } else {
    tab <- wb_initialise(tab, wb, ws_name)
  }

  tab$misc$ws_name <- ws_name

  tab

}
