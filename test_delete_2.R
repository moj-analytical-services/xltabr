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

xltabr:::combine_all_styles(tab)

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

openxlsx::openXL(tab$wb)
openxlsx::saveWorkbook(tab$wb, "testoutput.xlsx")

