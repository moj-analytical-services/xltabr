#' Get style sheet path
#'
#' Returns path to current style sheet which is set in global options. To change this path see set_style_path.
#' Default stylesheet used by package is [here](https://github.com/moj-analytical-services/xltabr/blob/dev/inst/extdata/styles.xlsx?raw=true)
#' @export
get_style_path <- function(){
  getOption("xltabr.style.path")
}

#' Get cell formats path
#'
#' Returns path to cell format definitions which is set in global options. To change this path see set_cell_format_path.
#' Default cell format used by package is [here](https://github.com/moj-analytical-services/xltabr/blob/dev/inst/extdata/style_to_excel_number_format.csv?raw=true)
#' @export
get_cell_format_path <- function(){
  getOption("xltabr.cell.format.path")
}

#' Get number formats path
#'
#' Returns path to number format definitions which is set in global options. To change this path see get_num_format_path.
#' Default number formats used by package is [here](https://github.com/moj-analytical-services/xltabr/blob/dev/inst/extdata/number_format_defaults.csv?raw=true)
#' @export
get_num_format_path <- function(){
  getOption("xltabr.number.format.path")
}

#' Set style sheet path
#'
#' Set the path to the style sheet to be used by package. To get this path see get_style_path.
#' Default cell formats used by xltabr is [here](https://github.com/moj-analytical-services/xltabr/blob/dev/inst/extdata/styles.xlsx?raw=true).
#' If no path is supplied the function sets the style sheet to default.
#' @param path the file path to the style sheet path (xlsx file). If NULL the function sets the cell format to the default option.
#' @export
set_style_path <- function(path = NULL){
  if(is.null(path)) path <- system.file("extdata", "styles.xlsx", package = "xltabr")
  if(file.exists(path)) options(list(xltabr.style.path = path)) else stop("Speficied file path does not exist")
}

#' Set cell format path
#'
#' Set the path to the cell formats to be used by xltabr. To get this path see get_cell_format_path.
#' Default cell format used by package is [here](https://github.com/moj-analytical-services/xltabr/blob/dev/inst/extdata/style_to_excel_number_format.csv?raw=true).
#' If no path is supplied the function sets the cell format to default to default.
#' @param path the file path to the cell formats (csv file). If NULL the function sets the cell format to the default option.
#' @export
set_cell_format_path <- function(path = NULL){
  if(is.null(path)) path <- system.file("extdata", "style_to_excel_number_format.csv", package = "xltabr")
  if(file.exists(path)) options(list(xltabr.cell.format.path = path)) else stop("Speficied file path does not exist")
}

#' Set number format path
#'
#' Set the path to the number formats to be used by xltabr. To get this path see get_num_format_path.
#' Default number format used by package is [here](https://github.com/moj-analytical-services/xltabr/blob/dev/inst/extdata/number_format_defaults.csv?raw=true).
#' If no path is supplied the function sets the cell format to default to default.
#' @param path the file path to the number formats (csv file). If NULL the function sets the nu,ber format to the default option.
#' @export
set_num_format_path <- function(path = NULL){
  if(is.null(path)) path <- system.file("extdata", "number_format_defaults.csv", package = "xltabr")
  if(file.exists(path)) options(list(xltabr.number.format.path = path)) else stop("Speficied file path does not exist")
}

get_cell_format_df <- function(){
  return(utils::read.csv(get_cell_format_path(), stringsAsFactors = FALSE, quote = "\"'"))
}
