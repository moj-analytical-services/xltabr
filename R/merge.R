#' Take a tab and merge the title rows
#'
#' @param tab The core tab object
#'
#' @export
#' @examples
#' crosstab <- read.csv(system.file("extdata", "example_crosstab.csv", package="xltabr"))
#' tab <- initialise()
#' tab <- add_body(tab, crosstab)
#' title_text <- c("Main title on first row", "subtitle on second row")
#' title_style_names <- c("title", "subtitle")
#' tab <- add_title(tab, title_text, title_style_names)
#' tab <- auto_merge_title_cells(tab)
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
#' @param tab The core tab object
#'
#' @export
#' @examples
#' crosstab <- read.csv(system.file("extdata", "example_crosstab.csv", package="xltabr"))
#' tab <- initialise()
#' tab <- add_body(tab, crosstab)
#' footer_text <- c("Footer contents 1", "Footer contents 2")
#' footer_style_names <- c("subtitle", "subtitle")
#' tab <- add_footer(tab, footer_text, footer_style_names)
#' tab <- auto_merge_footer_cells(tab)
auto_merge_footer_cells <- function(tab) {

  cols <- body_get_wb_cols(tab)
  rows <- footer_get_wb_rows(tab)

  for (r in rows) {
    openxlsx::mergeCells(tab$wb, tab$misc$ws_name, cols, r)
  }



  tab

}
