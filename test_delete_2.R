rm(list = ls())
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

# # # # # DEBUG
# tab <- xltabr:::initialise_debug()
#
# path <- system.file("extdata", "styles.xlsx", package = "xltabr")
#
# # if initialising style_catalogue do we want to reset it with line below?
# tab$style_catalogue <- list()
#
# wb <- openxlsx::loadWorkbook(path)
# openxlsx::readWorkbook(wb)
# #12
# i <- wb$styleObjects[[3]]
#
# r <- i$rows
# c <- i$cols
# cell <- openxlsx::readWorkbook(wb, rows = r, cols = c, colNames = FALSE, rowNames = FALSE)
#
# value <- cell[1, 1]
#
# tmp_list <- list()
# tmp_list$style <- i$style
#
# cell <- openxlsx::readWorkbook(wb, rows = r, cols = c + 1, colNames = FALSE, rowNames = FALSE)
# tmp_list$row_height <- cell[1, 1]
#
# sl <- xltabr:::convert_style_object(tmp_list)
# style_key <- xltabr:::create_style_key(sl)
# # # # #


tab <- xltabr::initialise() %>%
  xltabr::add_title(title_text, title_style_names) %>%
  xltabr::add_top_headers(h_list, row_style_names = c("body|indent_0", "body")) %>%
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

xltabr:::title_write_rows(tab)
xltabr:::top_headers_write_rows(tab)

tab <- xltabr:::add_styles_to_wb(tab, add_from = c("title","headers","body"))
openxlsx::saveWorkbook(tab$wb, "testoutput.xlsx")


### TESTING APPLYING STYLES
title_cells <- xltabr:::title_get_cell_styles_table(tab)
unique_styles <- title_cells %>%
  dplyr::select(style_name) %>%
  dplyr::distinct()
unique_title_styles <- as.character(unique_styles[,1])

table_cells <- xltabr:::top_headers_get_cell_styles_table(tab)
unique_styles <- table_cells %>%
  dplyr::select(style_name) %>%
  dplyr::distinct()
unique_cell_styles <- as.character(unique_styles[,1])

# take each cell style definition, convert it to a style_key and then add style_key to style_catalogue
for (style in unique_cell_styles){
  tab <- xltabr:::add_style_defintion_to_catelogue(tab, style)
}


# NEW BIT
sc <- tab$style_catalogue
sc_names <- names(sc)
names(sc) <- NULL
sc_values <- unlist(sc)

sc_table <- data.frame(style_name = sc_names, style_key = sc_values, stringsAsFactors = FALSE)

table_cells <- dplyr::left_join(table_cells, y = sc_table, by = "style_name")

unique_styles <- table_cells %>%
  dplyr::select(style_key) %>%
  dplyr::distinct()
unique_style_keys <- as.character(unique_styles[,1])

# get a unique set of style_keys

# Need to turn below into function once working
# First get a unique list of styles that are used in final workbook
style_key <- unique_style_keys[1]
for (sk in unique_style_keys){
  row_col_styles <- table_cells %>% dplyr::filter(style_key == sk)
  rows <- row_col_styles$row
  cols <- row_col_styles$col

  created_style <- xltabr:::convert_style_object(sk, convert_to_S4 = TRUE)
  openxlsx::addStyle(tab$wb, tab$wb$sheet_names[1], created_style, rows, cols)
}
# names(tab$style_catalogue)
# names(tab$style_dictionary)
#
openxlsx::saveWorkbook(tab$wb, "testoutput.xlsx")
