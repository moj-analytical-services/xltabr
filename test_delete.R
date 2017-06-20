library(magrittr)

title <- c("Table 1.1: Prison population by type of custody, age group and sex", "Subtitle text goes here", "")
title_style <- c("title", "subtitle", "spacer")

df <- readr::read_csv("xtab_example.csv")
df["type"] <- "Males and Females"
df <- dplyr::select(df, type, dplyr::everything())


cols <- c("type", "sentenced", "sentence_cat_1", "sentence_cat_2")


tab <- xltabr::initialise(df, column_headers = cols) %>%
       xltabr::add_titles(title, title_style) %>%
       xltabr::autoderive_formats_from_column_headers() %>%
       xltabr::combine_column_headers(cols) %>%
       xltabr::write_to_wb()

openxlsx::openXL(tab$wb)

