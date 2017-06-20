preprocess_df_orig_and_create_df_final <- function(tab) {
  tab$data$df_final <- tab$data$df_orig %>% dplyr::select(-dplyr::matches("meta_formatting_"))

  tab <- create_column_formats(tab)
  tab
}

# Use the column class to automatically guess the number formats we might want.
# TODO:  Make this a bit more sophisticated - e.g. do we want to introspect the column contents?
create_column_formats <- function(tab) {
  df <- tab$data$df_final
  names <- colnames(df)
  col_classes <- sapply(df, class)

  path <- system.file("extdata", "class_to_excel_number_format.csv", package = "xltabr")
  lookup <- readr::read_csv(path)

  lookup_list <- as.list(lookup$excel_format)
  names(lookup_list) <- lookup$class

  col_excel_number_format <- character()

  for (i in seq_along(names)) {

    name <- names[i]
    col_class <- col_classes[i]
    excel_format <- lookup_list[[col_class]]

    col_excel_number_format <- c(col_excel_number_format, excel_format)
  }

  attr(tab$data$df_final, "col_excel_number_format") <- col_excel_number_format

  tab

}
