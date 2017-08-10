# Need some unit tests that

library(magrittr)

# Test 1
ct <- reshape2::dcast(mtcars, am + gear ~ cyl, value.var= "mpg", margins=c("am", "gear"), fun.aggregate = mean)
headers <- colnames(ct)

tab <- xltabr::initialise()
tab <- xltabr::add_top_headers(tab, headers)
tab <- xltabr::add_body(tab, ct)
tab <- xltabr:::auto_detect_left_headers(tab)
tab <- xltabr:::auto_detect_body_title_level(tab)
tab <- xltabr:::auto_style_indent(tab)
tab <- xltabr::auto_style_number_formatting(tab)
tab <- xltabr:::write_all_elements_to_wb(tab)
tab <- xltabr:::add_styles_to_wb(tab)
openxlsx::openXL(tab$wb)

# Test 2
ct <- reshape2::dcast(mtcars, am + gear ~ cyl, value.var= "mpg", margins=c("am", "gear"), fun.aggregate = mean)
headers <- colnames(ct)

tab <- xltabr::initialise()
tab <- xltabr::add_top_headers(tab, headers)
tab <- xltabr::add_body(tab, ct)
tab <- xltabr::add_title(tab, "Here is a title")
tab <- xltabr:::auto_detect_left_headers(tab)
tab <- xltabr:::auto_detect_body_title_level(tab)
tab <- xltabr:::auto_style_indent(tab)
tab <- xltabr::auto_style_number_formatting(tab)
tab <- xltabr:::write_all_elements_to_wb(tab)
tab <- xltabr:::add_styles_to_wb(tab)
openxlsx::openXL(tab$wb)

# Test 3
ct <- reshape2::dcast(mtcars, am + gear ~ cyl, value.var= "mpg", margins=c("am", "gear"), fun.aggregate = mean)
tab <- xltabr::auto_crosstab_to_wb(ct, indent = TRUE, return_tab = TRUE)
openxlsx::openXL(tab$wb)
xltabr:::body_get_cell_styles_table(tab)

# Test 4
path <- system.file("extdata", "synthetic_data.csv", package="xltabr")
df <- read.csv(path, stringsAsFactors = FALSE)
ct <- reshape2::dcast(df, drive + age + colour + type ~ date, value.var= "value", margins=c("drive", "age", "colour", "type"), fun.aggregate = mean)
tab <- xltabr::auto_crosstab_to_wb(ct, indent = TRUE, return_tab = TRUE)
View(tail(tab$body$body_df))

openxlsx::openXL(tab$wb)
