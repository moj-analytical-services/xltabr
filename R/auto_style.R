# Functions in this file are used to attempt to auto-detect styles from the information in the table

#' Use the data type of the columns to choose an automatic Excel format for body cells
#'
#' This function reads currency styling from the styles defined [here](https://github.com/moj-analytical-services/xltabr/blob/master/inst/extdata/number_formats.xlsx?raw=true)
#' @param tab a table object
#' @param overrides a list containing any manual overrides where the user wants to provide their own style name
#' @export
auto_style_number_formatting <- function(tab, overrides = list()) {

  # Want to add number style to meta_col_ on body only -  tab$body$meta_col_
  col_classes <- sapply(tab$body$body_df_to_write, class)

  # This lookup table coverts
  path <- get_num_format_path()

  lookup_df <- utils::read.csv(path, stringsAsFactors = FALSE)
  # Convert to a named vector that can be used as a lookup
  lookup <- lookup_df$style_name
  names(lookup) <- lookup_df$class

  # Iterate through body columns applying lookup or override if exists
  body_cols <- names(tab$body$body_df_to_write)

  for (this_col_name in body_cols) {

    if (this_col_name %in% names(overrides)) {
      this_style <- overrides[[this_col_name]]
    } else {
      this_style <- lookup[col_classes[[this_col_name]][1]]
      if (is.na(this_style)) {
        stop(paste0("Trying to autoformat column of class ", col_classes[[this_col_name]], "but no style defined in number_format_defaults.csv"))
      }
    }

    if (is_null_or_blank(tab$body$meta_col_[this_col_name])) {
      tab$body$meta_col_[this_col_name] <- this_style
    } else {
      tab$body$meta_col_[this_col_name] <- paste(tab$body$meta_col_[this_col_name], lookup[this_col_name], sep = "|")
    }

  }

  tab

}


# Uses presence of '(all)' in the leftmost columns to detect the summary level
# of the row, and then writes this style information to tab$body$body_df
auto_style_body_rows <- function(tab, indent = FALSE, keyword = "(all)") {

  # If headers haven't been provided by the user, attempt to autodetect them
  if (is.null(tab$body$left_header_colnames)) {
    tab <- auto_detect_left_headers(tab)
  }

  # If no headers, return tab and message
  if (is.null(tab$body$left_header_colnames)) {
    message("No left header columns detected or provided")
    return(tab)
  }

  # Autodetect the 'summary level' e.g. header 1 is most prominent, header 2 next etc.
  tab <- auto_detect_body_title_level(tab, keyword)

  tab

}

#' Uses the presence of '(all)' in the leftmost columns of data to detect that these
#' columns are really left headers rather than body columns
#'
#' Populates tab$body$left_header_colnames automatically
#' @param tab a tab object
#' @param keyword The keyword to use to detect summarisation.  Uses '(all)' by default because this is what reshape2::dcast uses
#' @export
auto_detect_left_headers <- function(tab, keyword = "(all)") {
  # Looking to write tab$body$left_header_colnames

  # These must be character columns - stop if you hit a non character column

  # First find all leftmost character columns, then iterate from right to left, finding the first column with keyword in it.
  # This is the last left_header

  col_classes <- sapply(tab$body$body_df_to_write, class)

  if (!("character" %in% col_classes)) {
    tab$body$left_header_colnames = NULL
    return(tab)
  }

  rightmost_character <- min(which(col_classes != "character")) - 1

  if (rightmost_character == 0) {
    tab$body$left_header_colnames = NULL
    return(tab)
  }

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

# For title level
# Note we assume there are a maximum of 5 title levels
get_inv_title_count_title <- function(left_headers_df,
                                      keyword,
                                      allcount_to_level_translate = NULL) {
  # +-------+-------+-------+-----------+-------------+--------------+
  # | col1  | col2  | col3  | all_count | title_level | indent_level |
  # +-------+-------+-------+-----------+-------------+--------------+
  # | (all) | (all) | (all) |         3 | title_3     |              |
  # | (all) | (all) | -     |         2 | title_4     | indent_1     |
  # | (all) | -     | -     |         1 | title_5     | indent_2     |
  # | -     | -     | -     |         0 |             | indent_3     |
  # +-------+-------+-------+-----------+-------------+--------------+
  if (is.null(allcount_to_level_translate)) {
    allcount_to_level_translate = c("0" = NA, "1" = 5, "2" = 4, "3" = 3, "4" = 2, "5" = 1)
  }

  to_count <- (left_headers_df == keyword)
  all_count <- rowSums(to_count)

  #rows with higher (all) count should have a lower title value because title_1 is the most emphasized
  allcount_to_level_translate[as.character(all_count)]

}

# For indent level
get_inv_title_count_indent <- function(left_headers_df, keyword) {
  # +-------+-------+-------+-----------+-------------+--------------+
  # | col1  | col2  | col3  | all_count | title_level | indent_level |
  # +-------+-------+-------+-----------+-------------+--------------+
  # | (all) | (all) | (all) |         3 | title_3     |              |
  # | (all) | (all) | -     |         2 | title_4     | indent_1     |
  # | (all) | -     | -     |         1 | title_5     | indent_2     |
  # | -     | -     | -     |         0 |             | indent_3     |
  # +-------+-------+-------+-----------+-------------+--------------+

  to_count <- (left_headers_df == keyword)
  all_count <- rowSums(to_count)

  #rows with higher (all) count should have a lower title value because title_1 is the most emphasized
  all_count_inv <- max(all_count) - all_count
  all_count_inv[all_count_inv == 0] <- NA

  all_count_inv

}

#' Autodetect the 'title level' of each row in the cross tabulation
#' e.g. title 1 is most prominent, title 2 next etc.
#'
#' Uses the presence of '(all)' to detect the prominence.  The parameter allcount_to_level_translate allows the user to control how the count of '(all)' in the left header is translated into the header level
#' @param tab a tab object
#' @param keyword The keyword to use to detect summarisation.  Uses '(all)' by default because this is what reshape2::dcast uses
#' @param allcount_to_level_translate A named vector that provides a lookup - by default c("0" = NA, "1" = 5, "2" = 4, "3" = 3, "4" = 2, "5" = 1), which says that e.g. allcount 1 results in title_5 etc
#' @export
auto_detect_body_title_level <- function(tab, keyword = "(all)", allcount_to_level_translate = NULL) {

  # Stop if no left header colnames provided
  if (is.null(tab$body$left_header_colnames)) {
    return(tab)
  }

  left_headers_df <- tab$body$body_df_to_write[tab$body$left_header_colnames]

  all_count_inv <- get_inv_title_count_title(left_headers_df, keyword = keyword, allcount_to_level_translate = allcount_to_level_translate)

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

#' Consolidate the header columns into one, taking the rightmost value and applying indent
#'
#' e.g. a | b | (all) -> b
#' e.g. (all) | (all) | (all) -> Grand Total
#'
#' @param tab a tab object
#' @param keyword The keyword to use to detect summarisation.  Uses '(all)' by default because this is what reshape2::dcast uses
#' @param total_text The text to use for the grand total (a row where all the left headers are '(all)'. Defaults to Grand Total.
#' @param left_header_colname The column name of left header column, which is now a single column.
#'
#' @export

auto_style_indent <- function(tab, keyword = "(all)", total_text = NULL, left_header_colname = " ") {

  if (is.null(total_text)) {
    total_text = "Grand Total"
  }

  tab$misc$coalesce_left_header_colname = left_header_colname
  if (is.null(tab$body$left_header_colnames )) {
    message("You've called auto_style_indent, but there are no left_header_colnames to work with")
    return(tab)
  }

  left_headers_df <- tab$body$body_df_to_write[tab$body$left_header_colnames]

  orig_left_header_colnames <- tab$body$left_header_colnames

  # count '(all)'
  to_count <- (left_headers_df == keyword)
  all_count <- rowSums(to_count)

  # paste together all left headers
  concat <- do.call(paste, c(left_headers_df, sep = "=|="))

  # Split concatenated string into elements, and find last element that's not (all)
  elems <- strsplit(concat, "=\\|=", perl = TRUE)

  last_elem <- lapply(elems, function(x) {
    x <- x[x != keyword]
    if (length(x) == 0) {
      x <- total_text
    }
    utils::tail(x,1)
  })

  new_left_headers <- unlist(last_elem)

  # Remove original left_header_columns and replace with new
  cols <- !(names(tab$body$body_df_to_write) %in% tab$body$left_header_colnames)
  tab$body$body_df_to_write <- tab$body$body_df_to_write[cols]
  tab$body$body_df_to_write  <- cbind(new_left_headers = new_left_headers, tab$body$body_df_to_write, stringsAsFactors = FALSE)
  names(tab$body$body_df_to_write)[1] <- left_header_colname

  tab$body$left_header_colnames <- c(left_header_colname)

  # Now need to fix meta_left_header_row_ to include indents
  #Set meta_left_header_row_ to include relevant indents
  all_count_inv <- get_inv_title_count_indent(left_headers_df, keyword = keyword)

  col <- tab$body$body_df$meta_left_header_row_
  rows_indices_to_change <- not_na(all_count_inv)
  concat <- all_count_inv[rows_indices_to_change]
  concat <- paste0("indent_", concat)

  tab$body$body_df[rows_indices_to_change,"meta_left_header_row_"] <- paste(col[rows_indices_to_change], concat,sep = "|")

  # Update body$meta_col_
  tab$body$meta_col_ <- c(left_header_colname = tab$body$meta_col_[[1]], tab$body$meta_col_[cols])
  names(tab$body$meta_col_)[1] <- left_header_colname

  # Finally, if tab$top_headers has too many cols, remove the extra cols, removing from 2:n
  if (not_null(tab$top_headers$top_headers_list)) {
    len_th_cols <- length(tab$top_headers$top_headers_list[[1]])
    if (len_th_cols > length(colnames(tab$body$body_df_to_write))) {

      cols_to_delete <- 2:length(orig_left_header_colnames)
      cols_to_retain <- !(1:len_th_cols %in% cols_to_delete)

      tab$top_headers$top_headers_col_style_names <- tab$top_headers$top_headers_col_style_names[cols_to_retain]
      for (r in length(tab$top_headers$top_headers_list)) {

        tab$top_headers$top_headers_list[[r]] <- tab$top_headers$top_headers_list[[r]] [cols_to_retain]
        tab$top_headers$top_headers_list[[r]][1] <- left_header_colname
      }

    }
  }

  tab
}

# Add a right vertical border to the rightmost column of the left headers
add_left_header_vertical_border <- function(tab, stylename = "right_border") {

  # What's the right most cell in the
  right_most <- length(tab$body$left_header_colnames)

  tab$body$meta_col_[right_most] <- paste(tab$body$meta_col_[right_most], stylename, sep = "|")

  tab

}
