library(magrittr)

title_text <- c("Table 1.1: Prison population by type of custody, age group and sex", "Subtitle text goes here", "")
title_style_names <- c("title", "subtitle", "spacer")

footer_text <- c("Caveat 1 goes here", "Caveat 2 goes here")
footer_style_names <- c("footer", "footer")

r1 <- paste0("r1_col_", as.character(1:3))
r2 <- paste0("r2_col_", as.character(1:3))
h_list <- list(r1,  r2)

df <- data.frame(purrr::map(1:3, function(x) {1:4}))
colnames(df) <- letters[1:3]

tab <- xltabr::initialise() %>%
  xltabr::add_title(title_text, title_style_names) %>%
  xltabr::add_top_headers(h_list, col_style_names = c("c1", "c2", "c3"), row_style_names = c("r1", "r2")) %>%
  xltabr::add_body(df) %>%
  xltabr::add_footer(footer_text, footer_style_names)

xltabr:::title_write_rows(tab)
xltabr:::top_headers_write_rows(tab)
xltabr:::body_write_rows(tab)
xltabr:::footer_write_rows(tab)

openxlsx::openXL(tab$wb)


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


### TESTING APPLYING STYLES
title_cells <- xltabr:::title_get_cell_styles_table(tab)
unique_styles <- title_cells %>%
  dplyr::select(style_name) %>%
  dplyr::distinct()
unique_styles <- as.character(unique_styles[,1])


for (style in unique_styles){
  row_col_styles <- title_cells %>% dplyr::filter(style_name == style)
  rows <- row_col_styles$row
  cols <- row_col_styles$col

  openxlsx::addStyle(tab$wb, tab$wb$sheet_names[1], tab$style_catalogue[[tab$style_dictionary[[style]]]]$style, rows, cols)
}

names(tab$style_catalogue)
names(tab$style_dictionary)

openxlsx::saveWorkbook(tab$wb, "testoutput.xlsx")
