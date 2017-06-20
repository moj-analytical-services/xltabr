sheetExists <- function(wb, ws_name) {
  sheets <- openxlsx::sheets(wb)
  return(ws_name %in% sheets)
}

# Get number of columns in the data we've written to the sheet
df_width <- function(tab) {
  ncol(tab$data$df_final)
}

get_body_rows <- function(tab) {

  start <- tab$metadata$startcell_row + tab$metadata$rows_before_df_counter + tab$metadata$num_header_rows
  end <- start + nrow(tab$data$df_final) - 1

  start:end
}

get_body_columns <- function(tab) {
  start <- tab$metadata$startcell_column + tab$metadata$cols_before_df_counter + get_num_header_columns(tab)
  end <- start + ncol(tab$data$df_final) - get_num_header_columns(tab)

  start:end
}

get_all_columns <- function(tab) {
  start <- tab$metadata$startcell_column + tab$metadata$cols_before_df_counter
  end <- start + ncol(tab$data$df_final) - 1

  start:end

}

#todo
get_header_columns <- function(tab) {
  start <- tab$metadata$startcell_column + tab$metadata$cols_before_df_counter
  end <- start + get_num_header_columns(tab) -1

  start:end



}

get_num_header_columns <- function(tab) {
  length(tab$metadata$header_columns)
}

not_null <- function(x) {
  !(is.null(x))
}
