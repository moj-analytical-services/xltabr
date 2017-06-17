generate_df_final <- function(tab) {
  tab$data$df_final <- tab$data$df_orig %>% dplyr::select(-dplyr::matches("meta_formatting_"))
  tab
}
