# Need some unit tests that
open_output <- T
library(magrittr)

ct <- reshape2::dcast(mtcars, am + gear ~ cyl, value.var= "mpg", margins=c("am", "gear"), fun.aggregate = mean)
headers <- colnames(ct)
tab <- xltabr::initialise()
tab <- xltabr::add_top_headers(tab, headers)
tab <- xltabr::add_body(tab, ct)
tab <- xltabr:::auto_detect_left_headers(tab)
tab <- xltabr:::auto_detect_body_title_level(tab)
tab <- xltabr:::auto_style_indent(tab)
tab <- xltabr::auto_style_number_formatting(tab)
tab <- xltabr:::add_left_header_vertical_border(tab)
tab <- xltabr:::write_all_elements_to_wb(tab)
tab <- xltabr:::add_styles_to_wb(tab)
xltabr:::combine_all_styles(tab)

if (open_output) openxlsx::openXL(tab$wb) else {
  if (file.exists("test1.xlsx")) file.remove("test1.xlsx")
  openxlsx::saveWorkbook(tab$wb, "test1.xlsx")
}

tab$body$body_df
# Test 2
ct <- reshape2::dcast(mtcars, am + gear ~ cyl, value.var= "mpg", margins=c("am", "gear"), fun.aggregate = mean)
headers <- colnames(ct)
tab <- xltabr::initialise()
tab <- xltabr::add_top_headers(tab, headers)
tab <- xltabr::add_body(tab, ct)
tab <- xltabr::add_title(tab, "Here is a title")
tab <- xltabr::add_footer(tab, "Here is a footer")
tab <- xltabr:::auto_detect_left_headers(tab)
tab <- xltabr:::auto_detect_body_title_level(tab)
tab <- xltabr:::auto_style_indent(tab)
tab <- xltabr::auto_style_number_formatting(tab)
tab <- xltabr:::write_all_elements_to_wb(tab)
tab <- xltabr:::add_styles_to_wb(tab)
xltabr:::combine_all_styles(tab)

if (open_output) openxlsx::openXL(tab$wb) else {
  if (file.exists("test2.xlsx")) file.remove("test2.xlsx")
  openxlsx::saveWorkbook(tab$wb, "test2.xlsx")
}

# Test 3
ct <- reshape2::dcast(mtcars, am + gear ~ cyl, value.var= "mpg", margins=c("am", "gear"), fun.aggregate = mean)
tab <- xltabr::auto_crosstab_to_wb(ct, indent = TRUE, return_tab = TRUE)
if (open_output) openxlsx::openXL(tab$wb) else {
  if (file.exists("test3.xlsx")) file.remove("test3.xlsx")
  openxlsx::saveWorkbook(tab$wb, "test3.xlsx")
}
xltabr:::body_get_cell_styles_table(tab)

# Test 4
library(xltabr)
library(dplyr)

path <- system.file("extdata", "synthetic_data.csv", package="xltabr")
df <- read.csv(path, stringsAsFactors = FALSE)
ct <- reshape2::dcast(df, drive + age + colour + type ~ date, value.var= "value", margins=c("drive", "age", "colour", "type"), fun.aggregate = mean)
ct <- ct %>% dplyr::arrange(-row_number())


path <- system.file("extdata", "styles.xlsx", package = "xltabr")
cell_path <- system.file("extdata", "style_to_excel_number_format_alt.csv", package = "xltabr")
xltabr::set_style_path(path)
xltabr::set_cell_format_path(cell_path)

tab <- xltabr::auto_crosstab_to_wb(ct, indent = TRUE, return_tab = TRUE, titles = "This is the title")
openxlsx::openXL(tab$wb)

if (open_output) openxlsx::openXL(tab$wb) else {
  if (file.exists("test4.xlsx")) file.remove("test4.xlsx")
  openxlsx::saveWorkbook(tab$wb, "test4.xlsx")
}

tab$body$body_df$meta_left_header_row_

xltabr::set_style_path()
xltabr::set_cell_format_path()

# Test 5
path <- system.file("extdata", "synthetic_data.csv", package="xltabr")
df <- read.csv(path, stringsAsFactors = FALSE)
df$date <- as.Date(df$date)
ct <- reshape2::dcast(df, drive + age + colour ~ type, value.var= "date", margins=c("drive", "age", "colour"), fun.aggregate = min)
ct
tab <- xltabr::auto_crosstab_to_wb(ct, return_tab = TRUE)

View(tail(tab$body$body_df))

if (open_output) openxlsx::openXL(tab$wb) else {
  if (file.exists("test5.xlsx")) file.remove("test5.xlsx")
  openxlsx::saveWorkbook(tab$wb, "test5.xlsx")
}

# Test 6
path <- system.file("extdata", "test_autodetect.csv", package="xltabr")
df <- read.csv(path, stringsAsFactors = FALSE)
headers <- colnames(df)
tab <- xltabr::initialise()
tab <- xltabr::add_top_headers(tab, headers)
tab <- xltabr::add_body(tab, df)
tab <- xltabr:::auto_detect_left_headers(tab)
tab <- xltabr:::auto_detect_body_title_level(tab)
tab <- xltabr:::auto_style_indent(tab)
tab <- xltabr::auto_style_number_formatting(tab)
tab <- xltabr:::write_all_elements_to_wb(tab)
tab <- xltabr:::add_styles_to_wb(tab)
xltabr:::combine_all_styles(tab)

if (open_output) openxlsx::openXL(tab$wb) else {
  if (file.exists("test5.xlsx")) file.remove("test5.xlsx")
  openxlsx::saveWorkbook(tab$wb, "test5.xlsx")
}
xltabr:::combine_all_styles(tab)
df
