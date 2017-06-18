generate_df_final <- function(tab) {
  tab$data$df_final <- tab$data$df_orig %>% dplyr::select(-dplyr::matches("meta_formatting_"))
  tab <- apply_auto_classes(tab)
  tab
}


# apply_auto_classes <- function(tab) {
#   df <- tab$data$df_final
#   names <- colnames(df)
#   classes <- sapply(df, class)
#
#   path <- system.file("extdata", "class_to_excel_number_format.csv", package = "xltabr")
#   lookup <- readr::read_csv(path)
#
#   lookup_list <- as.list(lookup$excel_format)
#   names(lookup_list) <- lookup$class
#
#   for (i in seq_along(names)) {
#
#     name <- names[i]
#     class <- classes[i]
#     excel_format <- lookup_list[[class]]
#
#     class(tab$data$df_final[[name]]) <- excel_format
#   }
#
#
#   tab
#
# }
