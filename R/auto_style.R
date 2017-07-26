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
auto_style_body_rows <- function(tab, indent = FALSE, keyword = "(all)") {

  # If headers haven't been provided by the user, attempt to autodetect them
  if (is.null(tab$body$left_header_colnames)) {
    tab <- auto_detect_left_headers(tab)
  }

  # Autodetect the 'summary level' e.g. header 1 is most prominent, header 2 next etc.
  tab <- auto_detect_body_title_level(tab, keyword)


  tab

}

# Auto detect which of the columns are left headers
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

get_inv_title_count <- function(left_headers_df, keyword) {

  to_count <- (left_headers_df == keyword)
  all_count <- rowSums(to_count)

  #rows with higher (all) count should have a lower title value because title_1 is the most emphasized
  all_count_inv <- max(all_count) - all_count + 1
  all_count_inv[all_count == 0] <- NA

  all_count_inv

}

# Autodetect the 'title level' e.g. title 1 is most prominent, title 2 next etc.
auto_detect_body_title_level <- function(tab, keyword = "(all)") {

  left_headers_df <- tab$body$body_df_to_write[tab$body$left_header_colnames]

  all_count_inv <- get_inv_title_count(left_headers_df, keyword)

  # Append title level to both meta_row_ and meta_left_title_row_
  col <- tab$body$body_df$meta_row_[not_na(all_count_inv)]
  concat <- all_count_inv[not_na(all_count_inv)]
  concat <- paste0("title_", concat)
  tab$body$body_df[not_na(all_count_inv),"meta_row_"] <- paste(col, concat,sep = "|")

  col <- tab$body$body_df$meta_left_header_row_[not_na(all_count_inv)]
  concat <- all_count_inv[not_na(all_count_inv)]
  concat <- paste0("title_", concat)
  tab$body$body_df[not_na(all_count_inv),"meta_left_header_row_"] <- paste(col, concat,sep = "|")

  tab

}



# Consolidate the header columns into one, taking the rightmost value and applying indent
# e.g. a | b | (all) -> b
auto_style_indent(tab) {

  # scan from right to left finding first column that does not contain (all)
  left_headers_df <- tab$body$body_df_to_write[tab$body$left_header_colnames]




}
