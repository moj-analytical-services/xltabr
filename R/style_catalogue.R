# Xltabr uses the concept of an 'xltabr_style'.  This contains
# all the information needed to apply styles to cells
# An xltabr_style is a list with several keys, as follows
#   $style_list: An R list containing the properties that define the style, derived from openxlsx_style$style$as.list()
#   $s4style: An openxlsx style ready to be applied to some cells - this is an openxlsx_style$style
#   $style_string: A text string containing the name of the style e.g. body|indent_1
#   $row_height:  The row height of this style.
#   $excel_num_format:  The excel number format as a string

# Use .xlsx file containing styles to create the style catalogue, which is a list of xltabr_style s
style_catalogue_initialise <- function(tab) {

  path <- "styles.xlsx"
  num_path <- "style_to_excel_number_format.csv"
  xltabr_style_lookup <- convert_xlsx_and_csv_to_style_lookup()
  xltabr_style_lookup <- add_num_formats_to_style_lookup(xltabr_style_lookup)
  tab$style_catalogue <- xltabr_style_lookup
  tab
}

convert_xlsx_and_csv_to_style_lookup <- function() {

  # This is a bit tricky due to the potential for styles in the .xlsx file which
  # could have no styling ( e.g. body), and therefore do not exist in when we iterate through wb$styleObjects
  get_style_string_to_style_lookup <- function(wb) {

    # Create a list whose values are xltabr_styles.  Keys are cells references (row,column)
    xltabr_style_lookup <- list()
    for (openxlsx_style in wb$styleObjects) {
      for (row in openxlsx_style$rows) {
        for (col in openxlsx_style$cols) {
          key <-  paste(row,col, sep = ",")
          xltabr_style_lookup[[key]] <- list()
          xltabr_style_lookup[[key]]$style_list <- convert_s4style_to_style_list(openxlsx_style$style)
          xltabr_style_lookup[[key]]$s4style <- openxlsx_style$style
        }
      }
    }

    worksheet <- wb$worksheets[[1]] # We only expect styles in the first sheet of the .xlsx workbook
    # Associate a default style with rows in the xlsx that do not have a style.  Note at this point
    # we don't check for whether the contents of the cell are actually nothing
    for (row_index in worksheet$sheet_data$rows) {

      key <- paste(row_index,1, sep = ",")

      if (!key %in% names(xltabr_style_lookup)) {
        default_openxlsx_style <- openxlsx::createStyle()
        xltabr_style_lookup[[key]] <- list()
        xltabr_style_lookup[[key]]$style_list <- convert_s4style_to_style_list(default_openxlsx_style)
        xltabr_style_lookup[[key]]$s4style <- default_openxlsx_style
      }
    }

    # Finally convert the key of the lookup to the style name in the Excel cell rather than a row,column reference
    xltabr_style_lookup_new <- list()
    for (row_index in worksheet$sheet_data$rows) {
      old_key <- paste(row_index,1, sep = ",")
      value_sheet <- suppressWarnings(openxlsx::readWorkbook(wb, rows = row_index, cols = 1, colNames = FALSE))
      if (is_null_or_blank(value_sheet)) {
        next
      }
      new_key <- value_sheet[[1,1]]
      xltabr_style_lookup_new[[new_key]] <- xltabr_style_lookup[[old_key]]
      xltabr_style_lookup_new[[new_key]]$row_in_xlsx_stylesheet <- row_index
      xltabr_style_lookup_new[[new_key]]$style_string <- new_key
      xltabr_style_lookup_new[[new_key]]['excel_num_format'] <- list(NULL)
    }

    xltabr_style_lookup_new
  }

  # Row height is a property of the row, not in a style list, so needs to be dealt with separately
  add_row_height_to_style_string_lookup <- function(wb, xltabr_style_lookup) {

    new_xltabr_style_lookup <- list()

    for (xltabr_style in xltabr_style_lookup) {

      row <- as.character(xltabr_style$row_in_xlsx_stylesheet)
      row_height <- wb$rowHeights[[1]][row]  #This will return na if there's no entry

      if (is.na(row_height)) {
        row_height <- list(NULL)
      } else {
        row_height <- list(as.integer(row_height))
      }

      xltabr_style["row_height"] <- row_height  #must be [] due to non equivalence of l['a'] <- list(NULL) and l[['a']] <- list(NULL)
      new_xltabr_style_lookup[[xltabr_style$style_string]] <- xltabr_style
    }

    new_xltabr_style_lookup
  }

  wb <- openxlsx::loadWorkbook(get_style_path())
  xltabr_style_lookup <- get_style_string_to_style_lookup(wb)
  xltabr_style_lookup <- add_row_height_to_style_string_lookup(wb, xltabr_style_lookup)
  xltabr_style_lookup

}

add_num_formats_to_style_lookup <- function(xltabr_style_lookup) {

  num_formats_df <- utils::read.csv(get_cell_format_path(), stringsAsFactors = FALSE)

  # Convert dataframe into two vectors
  excel_formats <- num_formats_df$excel_format
  style_strings <- num_formats_df$style_name

  if (length(style_strings) == 0) {
    warning("Your number format excel contained no data.  For an example see xltabr/inst/extdata/style_to_excel_number_format.csv")
    return(NULL)
  }

  for (i in 1:length(style_strings)) {
    style_string <- style_strings[i]
    s4style <- openxlsx::createStyle(numFmt = excel_formats[[i]])

    xltabr_style_lookup[[style_string]] <- list()
    xltabr_style_lookup[[style_string]]$style_list <- convert_s4style_to_style_list(s4style)
    xltabr_style_lookup[[style_string]]$excel_num_format <- excel_formats[[i]]
    xltabr_style_lookup[[style_string]]$style_list$numFmt <- NULL
    xltabr_style_lookup[[style_string]]$s4style <- s4style
    xltabr_style_lookup[[style_string]]["row_height"] <- list(NULL)
    xltabr_style_lookup[[style_string]]$style_string <- style_string
  }

  xltabr_style_lookup

}

convert_s4style_to_style_list <- function(s4style) {

  style_list <- s4style$as.list()
  # For some reason fills do not export properly so workaround is added here
  if (any(c("fillFg", "fillBg") %in% names(style_list))) {
    style_list[["fill"]] <- list(fillFg = style_list[["fillFg"]], fillBg = style_list[["fillBg"]])
    style_list[["fillFg"]] <- NULL
    style_list[["fillBg"]] <- NULL
  }

  style_list

}

# Style string is like "body|header_1"
style_string_parse_and_combine <- function(style_lookup, style_string) {

  components <- strsplit(style_string, "\\|")[[1]]

  xltabr_styles_to_combine <- list()

  for (c in components) {
    xltabr_styles_to_combine[[c]] <- style_lookup[[c]]
  }

  # This is a reduce operation, but Reduce only works with vectors
  base_style <- xltabr_styles_to_combine[[1]]
  xltabr_styles_to_combine[[1]] <- NULL

  for (s in xltabr_styles_to_combine) {
    base_style <- style_inherit(base_style, s)
  }

  base_style

}

style_inherit <- function(base_xltabr_style, new_xltabr_style) {

  # Ignore fontColour = 1 (i.e. user just left it black)
  if (not_null(new_xltabr_style$style_list$fontColour)) {
    if (new_xltabr_style$style_list$fontColour == "1") {
      new_xltabr_style$style_list$fontColour <- base_xltabr_style$style_list$fontColour
    }
  }

  if (not_null(new_xltabr_style$style_list$fontName)) {
    tryCatch({
      base_xltabr_style$style_list$fontFamily <- NULL
      base_xltabr_style$style_list$fontScheme <- NULL
    })
  }

  new_style_list <- modifyList(base_xltabr_style$style_list, new_xltabr_style$style_list)

  # Inherit font decoration (we want bold + italic to become bold, italic, not italic)
  new_style_list$fontDecoration <- unique(c(base_xltabr_style$style_list$fontDecoration, new_xltabr_style$style_list$fontDecoration))



  new_xltabr_style$style_list <- new_style_list

  # Take the greater of the row heights if at least one is not null
  both_null <- all(is.null(c(base_xltabr_style$row_height, new_xltabr_style$row_height)))
  if (!both_null) {
    new_xltabr_style$row_height <- max(base_xltabr_style$row_height, new_xltabr_style$row_height)
  }

  # Inherit number format
  if (is.null(new_xltabr_style$excel_num_format)) {
    new_xltabr_style['excel_num_format'] <- list(base_xltabr_style$excel_num_format)
  }

  new_xltabr_style$s4style <- generate_s4_object_from_xltabr_style(new_xltabr_style)
  new_xltabr_style$style_string <- paste(base_xltabr_style$style_string, new_xltabr_style$style_string, sep = "|")

  new_xltabr_style

}

generate_s4_object_from_xltabr_style <- function(xltabr_style) {

  style_properties <- names(xltabr_style$style_list)

  if (is.null(xltabr_style$excel_num_format)) {
    excel_num_format <- "GENERAL"
  } else {
    excel_num_format <- xltabr_style$excel_num_format
  }

  out_style <- openxlsx::createStyle(numFmt = excel_num_format)

  for (prop in style_properties) {
    out_style[[prop]] <- xltabr_style$style_list[[prop]]
  }

  out_style
}

create_style_if_not_exists <- function(tab, style_string) {

  # If the style already exists
  if (style_string %in% names(tab$style_catalogue)) {
    return(tab)
  }

  # Otherwise create it
  new_xltabr_style <- style_string_parse_and_combine(tab$style_catalogue, style_string)
  tab$style_catalogue[[style_string]] <- new_xltabr_style
  tab

}

add_styles_to_wb <- function(tab) {

  #For each distinct style, see if it already exists in the style catalogue, otherwise,
  #create it, add it, and apply it

  full_table <- combine_all_styles(tab)

  unique_style_combinations <- unique(full_table$style_name)

  for (style_combination_string in unique_style_combinations) {

    row_col_styles <- full_table[full_table$style_name == style_combination_string,]

    #e.g. rows[1] cols[1] is the first cell this style will be applied to
    rows <- row_col_styles$row
    cols <- row_col_styles$col

    tab <- create_style_if_not_exists(tab, style_combination_string)

    xltabr_style <- tab$style_catalogue[[style_combination_string]]
    openxlsx::addStyle(tab$wb, tab$misc$ws_name, xltabr_style$s4style, rows, cols)

    tab <- update_row_heights(tab, rows, xltabr_style)

  }

  tab

}

update_row_heights <- function(tab, rows, xltabr_style) {
  # Row heights are a bit tricky because we need to iterate through each row
  # checking existing height and increasing height if existing height < new height
  # Need to add sheet in here probably - presumably [[1]] is the sheet
  worksheet_index <- which(tab$wb$sheet_names == tab$misc$ws_name)

  for (row in unique(rows)) {

    existing_row_height <-  get_prop_or_return_null(tab$wb$rowHeights[[worksheet_index]], as.character(row))
    new_row_height <- xltabr_style$row_height

    all_null <- all(is.null(c(existing_row_height, new_row_height)))

    if (!all_null) {
      new_row_height <- max(existing_row_height, new_row_height)
      openxlsx::setRowHeights(tab$wb, tab$misc$ws_name, row, new_row_height)
    }
  }

  tab

}

#' Manually add an Excel num format to the style catalogue
#'
#' @param tab a table object
#' @param style_string the name (key) in the tab$style_catalogue
#' @param excel_num_format an excel number format e.g. "#.00"
#'
#' @export
style_catalogue_add_excel_num_format <- function(tab, style_string, excel_num_format) {

  s4style <- openxlsx::createStyle(numFmt = excel_num_format)

  xltabr_style <- list()
  xltabr_style$s4style <- s4style
  xltabr_style$style_list <- convert_s4style_to_style_list(s4style)
  xltabr_style$excel_num_format <- excel_num_format
  xltabr_style['row_height'] <- list(NULL)

  tab$style_catalogue[[style_string]] <- xltabr_style

  tab
}


#' Manually add an openxlsx s4 style the style catalogue
#'
#' @param tab a table object
#' @param style_string the name (key) in the tab$style_catalogue
#' @param openxlsx_style an openxlsx s4 style
#' @param row_height the height of the row.  optional.
#'
#' @export
style_catalogue_add_openxlsx_style <- function(tab, style_string, openxlsx_style, row_height = NULL) {

  xltabr_style <- list()
  xltabr_style$s4style <- openxlsx_style
  xltabr_style$style_list <- convert_s4style_to_style_list(openxlsx_style)
  xltabr_style['row_height'] <- list(row_height)

  tab$style_catalogue[[style_string]] <- xltabr_style

  tab

}
