context("Test body")
test_that("column widths work as expected", {

  path <- system.file("extdata", "synthetic_data.csv", package = "xltabr")
  df <- read.csv(path, stringsAsFactors = FALSE)
  ct <- reshape2::dcast(df, drive + age + colour ~ type, value.var = "value", margins = c("drive", "age", "colour"), fun.aggregate = mean)
  tab <- xltabr::auto_crosstab_to_wb(ct, return_tab = TRUE, title = "Hello, I am a title - what what", fill_non_values_with = list(na = "**", nan = ".."))
  tab <- set_wb_widths(tab, left_header_col_widths = "auto", body_header_col_widths = c(7,14,28))

  # openxlsx::openXL(tab$wb)

  expect_error(set_wb_widths(tab, left_header_col_widths = "bob", body_header_col_widths = c(7,14,28)))
  expect_error(set_wb_widths(tab, left_header_col_widths = "auto", body_header_col_widths = c(7,14,28,8,60)))
})

test_that("Column widths work with offset", {

  ct <- reshape2::dcast(mtcars, am + gear ~ cyl, value.var = "mpg", margins = c("am", "gear"), fun.aggregate = mean)
  headers <- colnames(ct)

  tab <- xltabr::initialise(topleft_col = 4)
  tab <- xltabr::add_top_headers(tab, headers)
  tab <- xltabr::add_body(tab, ct)
  tab <- xltabr::add_title(tab, "Here is a title")
  tab <- xltabr::auto_detect_left_headers(tab)
  tab <- xltabr::auto_detect_body_title_level(tab)
  tab <- xltabr::auto_style_indent(tab)
  tab <- xltabr::auto_style_number_formatting(tab)
  tab <- xltabr::write_data_and_styles_to_wb(tab)
  tab <- xltabr::set_wb_widths(tab, left_header_col_widths = 20, body_header_col_widths = 10)
  # openxlsx::openXL(tab$wb)
})

