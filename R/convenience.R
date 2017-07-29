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
auto_df_to_xl <- function(df, auto_number_format = TRUE, titles = NULL, footers = NULL, auto_open = FALSE) {

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

  if (auto_open) {
    openxlsx::openXL(tab$wb)
  }

  tab <- xltabr:::add_styles_to_wb(tab)

  tab

}
