#' Create a new xltabr object for cross tabulation
#'
#' @param df The dataframe you want to output to Excel
#'
#' @return A list which contains the dataframe
#' @importFrom magrittr %>%
#' @name %>%#'
#' @export
initialise <- function(df, wb = openxlsx::createWorkbook(), ws_name = "Sheet1", startcell_row = 1, startcell_column = 1, column_headers = NULL) {
  tab <- list()
  tab$data <- list()
  tab$metadata <- list()
  tab$wb <- wb

  # Default values - functions are provided for the user to overwrite these
  tab$metadata$num_header_rows <- 1

  tab$metadata$header_columns <- column_headers


  tab$metadata$num_header_rows <- 1

  tab$metadata$startcell_row <- startcell_row
  tab$metadata$startcell_column <- startcell_column


  # Together the following counters keep track of how many rows of space this xltabr takes up
  # A counter which stores how many rows we've written before writing the core df
  tab$metadata$rows_before_df_counter <- 0

  # A counter which stores how many rows we've written before after the core df
  tab$metadata$cols_before_df_counter <- 0


  tab$style_catalogue <- list()
  tab <- add_default_styles(tab)

  tab$styles <- list()
  tab$styles$title_styles <- NULL

  tab$data$title <- character()
  tab$data$ws_name <- ws_name
  if (!(sheetExists(tab$wb, ws_name))) {
    openxlsx::addWorksheet(tab$wb, ws_name, gridLines = FALSE)
  }

  tab$data$df_orig <- df

  #
  tab <- preprocess_df_orig_and_create_df_final(tab)


  tab
}

#' @export
add_titles <- function(tab, title = "My title", title_style = "row_header_1") {
  tab$data$title <- title
  tab$data$title_style <- title_style
  tab
}


#' Guess row formats from presence of (all) in column headers
#'
#' @param tab The xtabr object
#'
#' @return tab
#' @export
autoderive_formats_from_column_headers <- function(tab, styles = NULL) {

  header_columns <- tab$metadata$header_columns
  if (is.null(styles)) {
    lookup <- dplyr::data_frame(indent = 0:10, meta_formatting_ = paste0("indent_", 0:10))
  } else {
    lookup <- dplyr::data_frame(indent = seq_along(styles, from=0), meta_formatting_ = paste0("indent_", seq_along(styles, from=0)))
  }

  tab$data$df_orig <- tab$data$df_orig %>%
    dplyr::mutate(indent = length(header_columns) - rowSums(.[header_columns] == "(all)") - 1) %>%
    dplyr::left_join(lookup, by="indent") %>%
    dplyr::select(-indent)

  tab
}



#' Convert multiple column headers in an xltabr object into
#'
#' @param df The dataframe you want to output to Excel
#'
#' @return A list which contains the dataframe
#' @export
combine_column_headers <- function(tab, columns) {

  # TODO:  This doesn't add a pad, it's to do with aligning right.  We just want to add spaces.

  cols_to_concat = lapply(as.list(columns), as.name)
  paste_fn <- dplyr::quos(paste(!!!cols_to_concat, sep="|'|"))

  get_last <- function(col) {
    elems <- strsplit(col, "\\|'\\|", perl=TRUE)
    last_elem <- lapply(elems, function(x) {
      x <- x[x != ""]
      if (length(x) == 0) {
        x <- ""
      }
      tail(x,1)
    })
    unlist(last_elem)
  }


  tab$data$df_final <- tab$data$df_final %>%
    dplyr::mutate(sum_margin_cols = rowSums(.[columns] == "(all)" )) %>%
    dplyr::mutate(header_column = !!! paste_fn) %>%
    dplyr::mutate(header_column = gsub("\\(all\\)","", header_column)) %>%
    dplyr::mutate(header_column = get_last(header_column)) %>%
    dplyr::select(-sum_margin_cols) %>%
    dplyr::select(.dots = -dplyr::one_of(columns)) %>%
    dplyr::select(header_column, dplyr::everything())


  tab$metadata$header_columns <- "header_column"
  tab <- create_column_formats(tab)

  tab
}
