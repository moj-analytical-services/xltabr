library(magrittr)

title_text <- c("Table 1.1: Prison population by type of custody, age group and sex", "Subtitle text goes here", "")
title_style_names <- c("title", "subtitle", "spacer")

r1 <- paste0("r1_col_", as.character(1:3))
r2 <- paste0("r2_col_", as.character(1:3))
h_list <- list( r1,  r2)

tab <- xltabr::initialise() %>%
  xltabr::add_title(title_text, title_style_names) %>%
  xltabr::add_top_headers(h_list, col_style_names = c("c1", "c2", "c3"), row_style_names = c("r1", "r2"))

xltabr:::title_write_rows(tab)
xltabr:::top_headers_write_rows(tab)

xltabr:::title_get_cell_styles_table(tab)
xltabr:::top_headers_get_cell_styles_table(tab)
openxlsx::saveWorkbook(tab$wb, "testoutput.xlsx")

