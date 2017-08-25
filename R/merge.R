#' Take a tab and merge the title rows
#'
#' @param df A tab
#'
#' @export
auto_merge_title_cells <- function(tab) {

  cols <- body_get_wb_cols(tab)
  rows <- title_get_wb_rows(tab)

  for (r in rows) {
    openxlsx::mergeCells(tab$wb, tab$misc$ws_name, cols, r)
  }

  tab

}

#' Take a tab and merge the footer rows
#'
#' @param df A tab
#'
#' @export
auto_merge_footer_cells <- function(tab) {

  cols <- body_get_wb_cols(tab)
  rows <- footer_get_wb_rows(tab)

  for (r in rows) {
    openxlsx::mergeCells(tab$wb, tab$misc$ws_name, cols, r)
  }



  tab

}
