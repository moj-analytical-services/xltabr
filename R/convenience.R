# These functions are convenience functinos that wrap other functionality to save the user typing

#' Take a data.frame in r and output an openxlsx wb objet
#'
#' @param df The data.frame to convert to Excel
#' @param auto_number_format Whether to automatically detect number format
#' @param title The title.  A character vector.  One element per row of title
#' @param footer Table footers.  A character vector.  One element per row of footer.
#' @param auto_open Automatically open Excel?
#'
#' @example auto_df_to_xl(mtcars, titles="the mtcars data", footers="note, this data is now quite old", auto_open=TRUE, auto_number_format = FALSE)
#' @export
auto_df_to_wb <- function(df, auto_number_format = TRUE, titles = NULL, footers = NULL, auto_open = FALSE, return_tab = FALSE, auto_merge = TRUE) {

  #Get headers from table
  headers <- names(df)

  tab <- xltabr::initialise() %>%
         xltabr::add_top_headers(headers) %>%
         xltabr::add_body(df)

  if (auto_number_format) {
    tab <- xltabr::auto_style_number_formatting(tab)
  }

  if (not_null(titles)) {
    tab <- xltabr::add_title(tab, titles)
  }

  if (not_null(footers)) {
    tab <- xltabr::add_footer(tab, footers)
  }

  tab <- write_all_elements_to_wb(tab)

  if (auto_merge) {
    tab <- xltabr::auto_merge_title_cells(tab)
    tab <- xltabr::auto_merge_footer_cells(tab)
  }

  tab <- xltabr:::add_styles_to_wb(tab)

  if (auto_open) {
    openxlsx::openXL(tab$wb)
  }


  if (return_tab) {
    return(tab)
  } else {
    return(tab$wb)
  }

}

#' Take a cross tabulation produced by `reshape2::dcast` and output a formatted openxlsx wb objet
#'
#' @param df A data.frame.  The cross tabulation to convert to Excel
#' @param top_headers. A list.  Custom top headers.
#' @param auto_number_format Whether to automatically detect number format
#' @param title The title.  A character vector.  One element per row of title
#' @param footer Table footers.  A character vector.  One element per row of footer.
#' @param auto_open Automatically open Excel?
#'
#' @example auto_df_to_xl(mtcars, titles="the mtcars data", footers="note, this data is now quite old", auto_open=TRUE, auto_number_format = FALSE)
#' @export
auto_crosstab_to_wb <- function(df,  auto_number_format = TRUE, top_headers = NULL, titles = NULL, footers = NULL, auto_open = FALSE, indent = TRUE, left_header_colnames = NULL,  vertical_border = TRUE, styles_xlsx = NULL, num_styles_csv = NULL, return_tab = FALSE, auto_merge = TRUE) {

  top_header_provided <- TRUE
  if (is.null(top_headers)) {
    top_header_provided <- FALSE
    top_headers  <- names(df)
  }

  tab <- xltabr::initialise(styles_xlsx = styles_xlsx, num_styles_csv = num_styles_csv) %>%
    xltabr::add_top_headers(top_headers) %>%
    xltabr::add_body(df,left_header_colnames = left_header_colnames)

  if (is.null(left_header_colnames)) {
    tab <- xltabr:::auto_detect_left_headers(tab)
  }

  tab <- xltabr:::auto_detect_body_title_level(tab)

  if (auto_number_format) {
    tab <- xltabr::auto_style_number_formatting(tab)
  }

  if (not_null(titles)) {
    tab <- xltabr::add_title(tab, titles)
  }

  if (indent) {
    tab <- xltabr:::auto_style_indent(tab)
  }

  if (vertical_border) {
    tab <- xltabr:::add_left_header_vertical_border(tab)
  }

  if (not_null(footers)) {
    tab <- xltabr::add_footer(tab, footers)
  }

  if (auto_merge) {
    tab <- xltabr::auto_merge_title_cells(tab)
    tab <- xltabr::auto_merge_footer_cells(tab)
  }

  tab <- write_all_elements_to_wb(tab)

  if (auto_open) {
    openxlsx::openXL(tab$wb)
  }

  tab <- xltabr:::add_styles_to_wb(tab)

  if (return_tab) {
    return(tab)
  } else {
    return(tab$wb)
  }


}
