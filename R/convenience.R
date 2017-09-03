# These functions are convenience functinos that wrap other functionality to save the user typing

#' Take a data.frame in r and output an openxlsx wb objet
#'
#' @param df The data.frame to convert to Excel
#' @param auto_number_format Boolean. Whether to automatically detect number formats of columns
#' @param title The title.  A character vector.  One element per row of title.
#' @param footers Table footers.  A character vector.  One element per row of footer.
#' @param left_header_colnames  The names of the columns that you want to designate as left headers
#' @param vertical_border Boolean. Do you want a left border?
#' @param auto_open Boolean. Automatically open Excel output.
#' @param return_tab  Boolean.  Return a tab object rather than a openxlsx workbook object
#' @param auto_merge Boolean.  Whether to merge cells in the title and footers to width of body
#' @param insert_below_tab. A existing tab object.  If provided, this table will be written on the same sheet, below the provided tab.
#'
#' @examples auto_df_to_wb(mtcars, titles="the mtcars data", footers="note, this data is now quite old", auto_number_format = FALSE)
#' @export
auto_df_to_wb <-
  function(df,
           auto_number_format = TRUE,
           titles = NULL,
           footers = NULL,
           left_header_colnames = NULL,
           vertical_border = TRUE,
           auto_open = FALSE,
           return_tab = FALSE,
           auto_merge = TRUE,
           insert_below_tab = NULL
           ) {

  #Get headers from table
  headers <- names(df)

  tab <- initialise(insert_below_tab = insert_below_tab) %>%
    add_top_headers(headers) %>%
    add_body(df,left_header_colnames = left_header_colnames)

  if (is.null(left_header_colnames)) {
    tab <- auto_detect_left_headers(tab)
  }

  if (auto_number_format) {
    tab <- auto_style_number_formatting(tab)
  }

  if (not_null(titles)) {
    tab <- add_title(tab, titles)
  }

  if (not_null(footers)) {
    tab <- add_footer(tab, footers)
  }

  tab <- write_all_elements_to_wb(tab)

  if (auto_merge) {
    tab <- auto_merge_title_cells(tab)
    tab <- auto_merge_footer_cells(tab)
  }

  if (vertical_border) {
    tab <- add_left_header_vertical_border(tab)
  }

  tab <- add_styles_to_wb(tab)

  if (auto_open) {
    openxlsx::openXL(tab$wb)
  }


  if (return_tab) {
    return(tab)
  } else {
    return(tab$wb)
  }

}

#' Take a cross tabulation produced by `reshape2::dcast` and output a formatted openxlsx wb object
#'
#' @param df A data.frame.  The cross tabulation to convert to Excel
#' @param auto_number_format Whether to automatically detect number format
#' @param top_headers A list.  Custom top headers. See [add_top_headers()]
#' @param titles The title.  A character vector.  One element per row of title
#' @param footers Table footers.  A character vector.  One element per row of footer.
#' @param auto_open Boolean. Automatically open Excel output.
#' @param indent Automatically detect level of indentation of each row
#' @param left_header_colnames  The names of the columns that you want to designate as left headers
#' @param vertical_border Boolean. Do you want a left border?
#' @param styles_xlsx File path (string).  If provided, the styles defined in this xlsx are used rather than the default. See [here](https://github.com/moj-analytical-services/xltabr/blob/dev/inst/extdata/styles.xlsx) for template.
#' @param num_styles_csv  File path.  If provided, overrides the default number styles, which can be found [here](https://github.com/moj-analytical-services/xltabr/blob/dev/inst/extdata/style_to_excel_number_format.csv).
#' @param return_tab  Boolean.  Return a tab object rather than a openxlsx workbook object
#' @param auto_merge Boolean.  Whether to merge cells in the title and footers to width of body
#' @param insert_below_tab A existing tab object.  If provided, this table will be written on the same sheet, below the provided tab.
#' @param total_text  The text that is used for the 'grand total' of a cross tabulation
#' @param include_header_rows  Boolean - whether to include or omit the header rows
#' @param number_format_overrides e.g. list("colname1" = "currency1") see [auto_style_number_formatting]
#' @param wb A existing openxlsx workbook.  If not provided, a new one will be created
#' @param ws_name The name of the worksheet you want to write to
#'
#' @export
auto_crosstab_to_wb <-
  function(df,
           auto_number_format = TRUE,
           top_headers = NULL,
           titles = NULL,
           footers = NULL,
           auto_open = FALSE,
           indent = TRUE,
           left_header_colnames = NULL,
           vertical_border = TRUE,
           styles_xlsx = NULL,
           num_styles_csv = NULL,
           return_tab = FALSE,
           auto_merge = TRUE,
           insert_below_tab = NULL,
           total_text = NULL,
           include_header_rows = TRUE,
           wb = NULL,
           ws_name = NULL,
           number_format_overrides = list()) {

  top_header_provided <- TRUE
  if (is.null(top_headers)) {
    top_header_provided <- FALSE
    top_headers  <- names(df)
  }


  tab <- initialise(styles_xlsx = styles_xlsx, num_styles_csv = num_styles_csv, insert_below_tab = insert_below_tab, wb = wb, ws_name = ws_name)

  if (include_header_rows) {
    tab <- add_top_headers(tab, top_headers)
  }

  tab <- add_body(tab, df,left_header_colnames = left_header_colnames)

  if (is.null(left_header_colnames)) {
    tab <- auto_detect_left_headers(tab)
  }

  tab <- auto_detect_body_title_level(tab)

  if (auto_number_format) {
    tab <- auto_style_number_formatting(tab, overrides = number_format_overrides)
  }

  if (not_null(titles)) {
    tab <- add_title(tab, titles)
  }

  if (indent) {
    tab <- auto_style_indent(tab, total_text = total_text)
  }

  if (vertical_border) {
    tab <- add_left_header_vertical_border(tab)
  }

  if (not_null(footers)) {
    tab <- add_footer(tab, footers)
  }

  if (auto_merge) {
    tab <- auto_merge_title_cells(tab)
    tab <- auto_merge_footer_cells(tab)
  }

  tab <- write_all_elements_to_wb(tab)

  if (auto_open) {
    openxlsx::openXL(tab$wb)
  }

  tab <- add_styles_to_wb(tab)

  if (return_tab) {
    return(tab)
  } else {
    return(tab$wb)
  }


}
