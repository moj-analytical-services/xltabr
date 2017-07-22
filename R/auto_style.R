# Functions in this file are used to attempt to auto-detect styles from the information in the table

#' Use the data type of the columns to choose an automatic Excel format
#' Number formats are only applied to body cells
#'
#' Note this reads currency styling from the styles defined [here](https://github.com/moj-analytical-services/xltabr/blob/dev/inst/extdata/number_formats.xlsx?raw=true)
#' @param tab a table object
#' @param overrides a list containing any manual overrides where the user wants to provide their own style name
#' @example auto_style_number_formatting(overrides = list("colname1" = "currency1"))
#' @export
auto_style_number_formatting <- function(tab, overrides = list()) {



  # Want to add number style to meta_col_ on body only -  tab$body$meta_col_
  col_classes <- sapply(tab$body$body_df_to_write, class)

  # This lookup table coverts
  path <- system.file("extdata", "class_to_excel_number_format.csv", package = "xltabr" )

  lookup_df <- read.csv(path, stringsAsFactors = FALSE)
  # Convert to a named vector that can be used as a lookup
  lookup <- lookup_df$excel_format
  names(lookup) <- lookup_df$class

  # Iterate through body columns applying lookup or override if exists
  body_cols <- names(tab$body$body_df_to_write)

  for (this_col_name in body_cols) {

    if (this_col_name %in% names(overrides)) {
      this_style <- overrides[[this_col_name]]
    } else {
      this_style <- lookup[col_classes[this_col_name]]
    }


    if (is_null_or_blank(tab$body$meta_col_[this_col_name])) {
      tab$body$meta_col_[this_col_name] <- this_style
    } else {
      tab$body$meta_col_[this_col_name] <- paste(tab$body$meta_col_[this_col_name], lookup[this_col_name], sep = "|")
    }

  }

  tab

}


# Uses the presence of '(all)' in the leftmost columns of data to detect that these
# columns are really left headers rather than body colummns
# auto_style_left_headers <- function(tab) {
#
#
#
#   ab$body$body_df_to_write

  # Work from the left of the table finding columns

  # if (is.null(styles)) {
  #   lookup <- dplyr::data_frame(indent = 0:10, meta_formatting_ = paste0("indent_", 0:10))
  # } else {
  #   lookup <- dplyr::data_frame(indent = seq_along(styles, from=0), meta_formatting_ = paste0("indent_", seq_along(styles, from=0)))
  # }

  # tab$data$df_orig <- tab$data$df_orig %>%
  #   dplyr::mutate(indent = length(header_columns) - rowSums(.[header_columns] == "(all)") - 1) %>%
  #   dplyr::left_join(lookup, by="indent") %>%
  #   dplyr::select(-indent)

#   tab
#
# }

auto_detect_left_headers <- function(tab, keyword = "(all)") {
  # Looking to write tab$body$left_header_colnames

  # These must be character columns - stop if you hit a non character column

  # First find all leftmost character columns, then iterate from right to left, finding the first column with keyword in it.
  # This is the last left_header

  col_classes <- sapply(tab$body$body_df_to_write, class)

  rightmost_character <- min(which(col_classes != "character")) -1
  rightmost_character_cols <- rightmost_character:1

  found <- FALSE
  for (col in rightmost_character_cols) {
    this_col <- tab$body$body_df_to_write[,col]

    if (any(keyword == this_col)) {
      found <- TRUE
      break
    }
  }

  # TODO: this is not robust when multiple columns have the same name
  if (found) {
    left_col_names <- names(tab$body$body_df_to_write)[1:col]
  } else {
    left_col_names <- NULL
  }

  tab$body$left_header_colnames <- left_col_names

  tab



}

# auto_style_based_on_number_of_all() {
#
# }
