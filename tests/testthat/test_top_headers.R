context("Test top headers")
library(magrittr)

test_that("Test top headers integrety preserved when we coalesce left header columns", {

  path <- system.file("extdata", "synthetic_data.csv", package="xltabr")
  df <- read.csv(path, stringsAsFactors = FALSE)

  ct <- reshape2::dcast(df, drive + age + colour ~ type, value.var= "value", margins=c("drive", "age", "colour"), fun.aggregate = mean)

  # tab <- xltabr::auto_crosstab_to_wb(ct, return_tab = TRUE)
  headers <- colnames(ct)

  tab <- xltabr::initialise()
  tab <- xltabr::add_top_headers(tab, headers)
  tab <- xltabr::add_body(tab, ct)
  tab <- xltabr:::auto_detect_left_headers(tab)
  tab <- xltabr:::auto_detect_body_title_level(tab)
  tab <- xltabr:::auto_style_indent(tab)

  # 'new left headers' should read ''.  Need optional argument on 'auto style indent' which lets user set the text in top left.
  t1 = all(tab$top_headers$top_headers_list[[1]] == c(" ", "Sedan", "Sport", "Supermini"))

  testthat::expect_true(t1)


})

# Add test where multiple columns have the same name.  How important is this ?
