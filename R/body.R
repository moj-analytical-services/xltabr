#' Initialise the body element of the tab, creating all the required properties
#' the core tab object
#'
#' @param tab The core tab object
body_initialise <- function(tab) {

  tab$body <- list()

  # body_df contains the original data which was passed to to body, augmented
  # with two additional columns that store styling information, called
  # meta_row_ and meta_left_header_row_.  We create a duplicate that
  # contains only the data that we want to write to the worksheet, called body_df_to_write
  # so we don't accidentally write the meta data to the Excel worksheet
  tab$body$body_df <- NULL  # This will include a column for meta_row_ and meta_left_header_row_
  tab$body$body_df_to_write <- NULL # This is the df that's actually written to wb
  tab$body$left_header_colnames <- NULL # This will contain the column names for each of the columns which are left headers

  tab
}


#' Add a table body (a df) to the tab.
#'
#' @param tab The core tab object
#' @param df A data frame containing the data you want to write to write to Excel
#' @param left_header_colnames The names of the columns in the df which are left headers, as opposed to the main body
#' @param row_style_names Manually specify the styles that apply to each row (as opposed to using the autodetect functions).  Styles provided must be present in the style catalogue
#' @param left_header_style_names Manually specify the addition styles that will apply to cells in the left headers.  Must be present in styles catalogue
#' @param col_style_names Manually specify the additional styles that will be applied to each column.  Must be present in styles catalogue
#' @param fill_non_values_with Manually specify a list of strings that will replace non numbers types NA, NaN, Inf and -Inf. e.g. list(na = '*', nan = '', inf = '-', neg_inf = '--'). Note: NaNs are not treated as NAs.
#'
#' @export
add_body <-
  function(tab,
           df,
           left_header_colnames = NULL,
           row_style_names = NULL,
           left_header_style_names = NULL,
           col_style_names = NULL,
           fill_non_values_with = list(na = NULL, nan = NULL, inf = NULL, neg_inf = NULL)) {

  # Make all factors character
  i <- sapply(df, is.factor)
  df[i] <- lapply(df[i], as.character)

  tab$body$body_df <- df
  tab$body$body_df_to_write <- df
  tab$body$left_header_colnames <- left_header_colnames

  # Establish 'meta' columns, which contain style names for each row (and will not be written to the wb)
  tab$body$body_df$meta_row_ <- "body" # Fill whole col
  tab$body$body_df$meta_left_header_row_ <- "body|left_header" #Fill whole col

  # The meta_ columns deal with styling that varies by row.
  # We also need to set up a vector that contains a style_name per column
  n <- names(tab$body$body_df_to_write)
  tab$body$meta_col_ <- rep("", length(n))
  names(tab$body$meta_col_) <- n
  # If the user has provided row or col style names, add them.  Use recycling rules.
  if (not_null(row_style_names)) {
    m <- rbind(tab$body$body_df[[1]], row_style_names)
    row_style_names <- m[2,]
    tab$body$body_df$meta_row_ <- row_style_names
    tab$body$user_provided_row_style_names <- row_style_names
  }

  if (not_null(left_header_style_names)) {
    m <- rbind(tab$body$body_df[[1]], left_header_style_names)
    left_header_style_names <- m[2,]
    tab$body$body_df$meta_left_header_row_ <- left_header_style_names
    tab$body$user_provided_left_header_style_names <- left_header_style_names
  }

  if (not_null(col_style_names)) {
    m <- rbind(names(df), col_style_names)
    col_style_names <- m[2,]
    tab$body$meta_col_ <- col_style_names
    tab$body$user_provided_col_style_names <- col_style_names
  }

  # Run checks on fill_non_values_with
  possible_names <- c('na', 'nan', 'inf', 'neg_inf')

  if (any(!names(fill_non_values_with) %in% possible_names) | max(unlist(lapply(fill_non_values_with, length))) > 1 | !is.list(fill_non_values_with)) {
    stop(paste0('Error: The input fill_non_values_with must be a list and cannot have names other than: ', paste0(possible_names, collapse = ", "), '. Each element in this list must also be a single value (e.g. not a vector or list)'))
  }
  tab$body$fill_non_values_with <- fill_non_values_with

  tab
}

#' Compute the columns of the workbook which are occupied by the body
#'
#' @param tab The core tab object
body_get_wb_cols <- function(tab) {

  if (is.null(tab$body$body_df)) {
    return(NULL)
  }

  tlc <- tab$extent$topleft_col

  cols_vec <- seq_along(tab$body$body_df_to_write)

  wb_cols <- seq_along(cols_vec) + tlc - 1

  wb_cols
}

#' Compute the columns of the workbook which are occupied by the left header columns (which are a subset of the body columns)
#'
#' @param tab The core tab object
body_get_wb_left_header_cols <- function(tab){

  if (is.null(tab$body$body_df)) {
    return(NULL)
  }

  tlc <- tab$extent$topleft_col

  cols_vec <- seq_along(tab$body$left_header_colnames)

  wb_cols <- seq_along(cols_vec) + tlc - 1

  wb_cols
}

#' Compute the rows of the workbook which are occupied by the body
#'
#' @param tab The core tab object
body_get_wb_rows <- function(tab) {
  if (is.null(tab$body$body_df)) {
    return(NULL)
  }
  offset <- top_headers_get_bottom_wb_row(tab)
  seq_along(tab$body$body_df[[1]]) + offset
}



#' Compute the bottom (lowest) row of the workbook occupied by the body
#' If the body does not exist, returns the last row of the previous element (the top header)
#'
#' @param tab The core tab object
body_get_bottom_wb_row <- function(tab) {
  top_headers_bottom <- top_headers_get_bottom_wb_row(tab)
  b_rows <- body_get_wb_rows(tab)

  max(c(top_headers_bottom, b_rows))
}

#' Compute the rightmost column of the workbook which is occupied by the body
#'
#' @param tab The core tab object
body_get_rightmost_wb_col <- function(tab) {
  rightmost_th <- top_headers_get_rightmost_wb_col(tab)
  b_cols <- body_get_wb_cols(tab)

  max(c(rightmost_th, b_cols))
}

#' Create table with columns |row|col|style name| containing the styles names for each cell of the body
#'
#' @param tab The core tab object
body_get_cell_styles_table <- function(tab) {

  # Approach is to start by creating a table of |row|col|body_style|left_header_style|top_header_style
  # See https://www.draw.io/#G0BwYwuy7YhhdxY2hGQnVGNFN6QkE

  rows <- body_get_wb_rows(tab)

  if (length(rows) == 0) {
    df <- data.frame("row" = integer(0), "col" = integer(0), "style_name" = character(0))
    return(df)
  }

  r <- body_get_wb_rows(tab)
  c_all_body <- body_get_wb_cols(tab)
  #Left headers only
  lh_cols <- body_get_wb_left_header_cols(tab)

  #c_all_body is all body cells - for body styling we need to remove header rows
  c_body_no_lh <- c_all_body[!(c_all_body %in% lh_cols)]

  df <- expand.grid(row = r, col = c_body_no_lh)

  #All cells get body
  df_br <- data.frame(row = r, style_name = tab$body$body_df$meta_row_, stringsAsFactors = FALSE) #br stands for body row
  df <- merge(df, df_br, by = "row") #At this point, df has all row col combos for the body, but not the left columns, and the given row style

  #If left header columns actually exist
  if (length(lh_cols) > 0) {
    #a table containing each row and its associated style for the header rows
    df_lhc_r <- data.frame(row = r, style_name = tab$body$body_df$meta_left_header_row_, stringsAsFactors = FALSE)

    hcs <- data.frame(col = lh_cols, stringsAsFactors = FALSE)
    df_hcs <- merge(df_lhc_r, hcs) # a df that contains |rowcol|left_header style for all cols and rows of left headers,

    df <- rbind(df, df_hcs)
  }

  #Add a final column that includes the column style information - i.e. top header styles
  df_th <- data.frame(col = c_all_body, top_header_style = tab$body$meta_col_, stringsAsFactors = FALSE)
  df <- merge(df, df_th, by = "col")

  df$style_name <- paste(df$style_name, df$top_header_style, sep = "|")
  df <- df[, c("row", "col", "style_name")]
  df$style_name <- remove_leading_trailing_pipe(df$style_name)

  df
}

#' Write all the body data to the workbook (but do not write style information)
#'
#' @param tab The core tab object
body_write_rows <- function(tab) {

  if (is.null(tab$body$body_df_to_write)) {
    return(tab)
  }

    ws_name <- tab$misc$ws_name

    data <- tab$body$body_df_to_write

    col <- min(body_get_wb_cols(tab))
    row <- min(body_get_wb_rows(tab))

    openxlsx::writeData(tab$wb, ws_name, data, startRow = row, startCol = col, colNames = FALSE)

    # Replace non-values with specified values
    if (not_null(tab$body$fill_non_values_with$na)) {
      row_values <- body_get_wb_rows(tab)
        for (c in 1:ncol(data)) {
          for (r in row_values[is_truely_na(data[,c])]) openxlsx::writeData(tab$wb, ws_name, tab$body$fill_non_values_with$na, startRow = r, startCol = c, colNames = FALSE)
        }
    }

    if (not_null(tab$body$fill_non_values_with$nan)) {
      row_values <- body_get_wb_rows(tab)
      for (c in 1:ncol(data)) {
        for (r in row_values[is.nan(data[,c])]) openxlsx::writeData(tab$wb, ws_name, tab$body$fill_non_values_with$nan, startRow = r, startCol = c, colNames = FALSE)
      }
    }

    if (not_null(tab$body$fill_non_values_with$inf)) {
      row_values <- body_get_wb_rows(tab)
      for (c in 1:ncol(data)) {
        for (r in row_values[is.infinite(data[,c]) & data[,c] > 0]) openxlsx::writeData(tab$wb, ws_name, tab$body$fill_non_values_with$inf, startRow = r, startCol = c, colNames = FALSE)
      }
    }

    if (not_null(tab$body$fill_non_values_with$neg_inf)) {
      row_values <- body_get_wb_rows(tab)
      for (c in 1:ncol(data)) {
        for (r in row_values[is.infinite(data[,c]) & data[,c] < 0]) openxlsx::writeData(tab$wb, ws_name, tab$body$fill_non_values_with$neg_inf, startRow = r, startCol = c, colNames = FALSE)
      }
    }
    tab
}

#' Set column widths to tab workbook
#'
#' @param tab The core tab object
#' @param left_header_col_widths Width of row header columns you wish to set in Excel column width units. If singular, value is applied to all row header columns. If a vector, vector must have length equal to the number of row headers in workbook. Use special case "auto" for automatic sizing. Default (NULL) leaves column widths unchanged.
#' @param body_header_col_widths Width of body header columns you wish to set in Excel column width units. If singular, value is applied to all body columns. If a vector, vector must have length equal to the number of body headers in workbook. Use special case "auto" for automatic sizing. Default (NULL) leaves column widths unchanged.
#'
#' @export
set_wb_widths <- function(tab, left_header_col_widths = NULL, body_header_col_widths = NULL){

  ws_name <- tab$misc$ws_name

  # get row header columns
  rhc <- body_get_wb_left_header_cols(tab)

  # calculate body columns
  all_cols <- body_get_wb_cols(tab)
  bc <- all_cols[!all_cols %in% rhc]

  # Set row hearder column widths
  if (not_null(left_header_col_widths)) {
    # Error checking
    if (typeof(left_header_col_widths) == "character") {
      if (left_header_col_widths != 'auto') stop('If left_header_col_widths is a character it must be set to \"auto\". Otherwise it must be a number or vector of numbers')
    } else {
      if (!typeof(left_header_col_widths) %in% c('double','integer')) stop('If left_header_col_widths is a character it must be set to \"auto\". Otherwise it must be a number or vector of numbers')
    }
    if (length(left_header_col_widths) != 1 & length(left_header_col_widths) != length(rhc)) {
      stop(paste0('left_header_col_widths must either be a single value or a vector equal to the number of row_header_columns (', length(rhc),')'))
    }

    openxlsx::setColWidths(tab$wb, sheet = ws_name, cols = rhc, widths = left_header_col_widths, ignoreMergedCells = T)
  }

  # Set body hearder column widths
  if (not_null(body_header_col_widths)) {
    # Error checking
    if (typeof(body_header_col_widths) == "character") {
      if (body_header_col_widths != 'auto') stop('If body_header_col_widths is a character it must be set to \"auto\". Otherwise it must be a number or vector of numbers')
    } else {
      if (!typeof(body_header_col_widths) %in% c('double','integer')) stop('If body_header_col_widths is a character it must be set to \"auto\". Otherwise it must be a number or vector of numbers')
    }
    if (length(body_header_col_widths) != 1 & length(body_header_col_widths) != length(bc)) {
      stop(paste0('body_header_col_widths must either be a single value or a vector equal to the number of row_header_columns (', length(bc),')'))
    }

    openxlsx::setColWidths(tab$wb, sheet = ws_name, cols = bc, widths = body_header_col_widths, ignoreMergedCells = T)
  }

  return(tab)
}
