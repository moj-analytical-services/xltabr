


apply_column_styles <- function(tab) {

  # Apply


  # TODO: Allow user to apply specific overrides here


}


apply_auto_classes <- function(tab) {
  df <- tab$data$df_final
  names <- colnames(df)
  classes <- sapply(df, class)

  path <- system.file("extdata", "class_to_excel_number_format.csv", package = "xltabr")
  lookup <- readr::read_csv(path)

  lookup_list <- as.list(lookup$excel_format)
  names(lookup_list) <- lookup$class

  for (i in seq_along(names)) {

    name <- names[i]
    class <- classes[i]
    excel_format <- lookup_list[[class]]

    # Go into the workbook and update the format of cells

    class(tab$data$df_final[[name]]) <- excel_format
  }


  tab

}


apply_row_styles <- function(tab) {


  start_row <- tab$metadata$startcell_row + tab$metadata$rows_before_df_counter
  start_column <- tab$metadata$startcell_column
  end_column <- start_column + ncol(tab$data$df_final) - 1

  ws_name <- tab$data$ws_name

  for (i in seq_along(tab$data$df_orig$meta_formatting_)) {
    write_row <- start_row + i
    style_name <- tab$data$df_orig$meta_formatting_[i]
    this_style <- tab$style_catalogue[[style_name]]$style
    row_height <- tab$style_catalogue[[style_name]]$row_height

    openxlsx::addStyle(tab$wb, ws_name, this_style, rows = write_row, cols=start_column:end_column)
    openxlsx::setRowHeights(tab$wb, ws_name, rows=write_row, heights=row_height)


  }

}



apply_styles_to_wb <- function(tab) {

  apply_row_styles(tab)

}
