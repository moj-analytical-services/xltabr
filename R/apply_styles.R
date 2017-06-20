apply_column_styles <- function(tab) {

  # Apply
  apply_col_number_formats(tab)

  # TODO: Allow user to apply specific overrides here


}


apply_col_number_formats <- function(tab) {

  #Need to iterate through all the cells of the body, column by columnm, updating styles

  #TODO:  This is currently broken because applying a style overrides previous formatting information
  df <- tab$data$df_final
  col_names <- colnames(df)
  formats <- attr(tab$data$df_final, "col_excel_number_format")

  all_cols <- get_all_columns(tab)

  # Iterate through the columns
  # TODO:  How do we make this faster???
  for (col_num in seq_along(all_cols)) {

    col_in_wb <- all_cols[col_num]
    excel_format <- formats[col_num]

    # For each cell in this column
    body_rows <- get_body_rows(tab)

    # Make new format with the style
    style_overrides <- openxlsx::createStyle(numFmt = excel_format)

    for (row in body_rows) {
      xltabr:::update_style_in_cell(tab,row, col_in_wb,style_overrides)
    }

  }

  tab

}


apply_row_styles <- function(tab) {

  ws_name <- tab$data$ws_name
  body_rows <- get_body_rows(tab)
  header_columns <- get_header_columns(tab)
  body_columns <- get_body_columns(tab)


  # For each row ihe main table
  for (i in seq_along(body_rows)) {
    write_row <- body_rows[i]
    style_name <- tab$data$df_orig$meta_formatting_[i]
    this_style <- tab$style_catalogue[[style_name]]$style
    row_height <- tab$style_catalogue[[style_name]]$row_height

    #Apply to header column cels
    openxlsx::addStyle(tab$wb, ws_name, this_style, rows = write_row, cols=header_columns)
    openxlsx::setRowHeights(tab$wb, ws_name, rows=write_row, heights=row_height)

    #Apply to body rows
    style_overrides <- openxlsx::createStyle(indent = 0)
    new_style <- update_style_object(this_style, style_overrides)
    openxlsx::addStyle(tab$wb, ws_name, new_style, rows = write_row, cols=get_body_columns(tab))

  }




}

apply_styles_to_wb <- function(tab) {

  apply_row_styles(tab)
  apply_column_styles(tab)

}


# Unfortunately it seems that finding the currently styling of a particular row and column is quite difficult
# Create a df with a list column that contains all the current styles for all the worksheets that can then easily be lookuped up
update_styles_lookup <- function(tab) {

  stobs <- tab$wb$styleObjects

  counter <- 1

  style <- list()
  col <- integer()
  row <- integer()

  for (stob in stobs) {
    for (i in seq_along(stob$rows)) {
      row[counter] <- stob$rows[i]
      col[counter] <- stob$cols[i]
      style[counter] <- stob$style
      counter = counter + 1
    }
  }

  tab$styles_lookup <- dplyr::data_frame(col = col, row = row, style = style)

  tab

}

get_style <- function(tab, row_find, col_find) {

  filtered <- tab$styles_lookup %>%
    dplyr::filter(row == get("row_find")) %>%
    dplyr::filter(col == get("col_find")) %>%
    dplyr::select(style)

  if (nrow(filtered) == 1 ) {
    return(filtered[[1,1]])
  } else {
    return(NULL)
  }



}

# Takes a style object and updates it based on a style
update_style_object <- function(style_old, style_overrides) {


  if (is.null(style_old)) {
    return(style_overrides)
  }

  if (is.null(style_overrides)) {
    return(style_old)
  }

  # styles_to_copy <-  c("valign", "borderLeft", "fontDecoration", "getClass", "borderTopColour", "indent", "borderLeftColour", "halign", "borderBottom", "borderRightColour", "fill", "initialize", "borderRight", "fontScheme", "numFmt", "wrapText", "textRotation", "xfId", "fontColour", "borderTop", "borderBottomColour", "show", "fontFamily", "fontSize", "fontName")
  styles_to_copy <- names(as.list(style_old))

  # Since the style is implemented as an s4 object, we first want to create a new object
  new_style <- openxlsx::createStyle()

  for (prop in styles_to_copy) {
    new_style[[prop]] <- style_old[[prop]]
  }

  for (prop in styles_to_copy) {
    if (not_null(style_overrides[[prop]])) {
      new_style[[prop]] <- style_overrides[[prop]]
    }
  }

  new_style

}

# Update the style in a cell r,c
update_style_in_cell <- function(tab, row, col, style_overrides) {
  # First get style from cell
  tab <- update_styles_lookup(tab)
  current_style <- get_style(tab, row, col)
  updated_style <- update_style_object(current_style, style_overrides)



  # Apply new style to cell
  openxlsx::addStyle(tab$wb, tab$data$ws_name, updated_style, rows = row, cols = col)
}

