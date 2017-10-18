# This script contains regression tests corresponding to issues in github

# https://github.com/moj-analytical-services/xltabr/issues/31
context("Test regressions")
library(magrittr)

test_that("test styles table does not have factors", {
  path <- system.file("extdata", "synthetic_data.csv", package = "xltabr")
  df <- read.csv(path, stringsAsFactors = FALSE)
  ct <- reshape2::dcast(df, drive + age + colour ~ type, value.var = "value", margins=c("drive", "age", "colour"), fun.aggregate = mean)
  tab <- xltabr::auto_crosstab_to_wb(ct, return_tab = TRUE, titles = "title", footers = "footers")

  df <- xltabr:::combine_all_styles(tab)
  t1 = (class(df$style_name) == "character")
  testthat::expect_true(t1)
})


test_that("test no warning are issued from style_catalogue_xlsx_import", {

  path <- system.file("extdata", "styles_pub.xlsx", package = "xltabr")
  xltabr::set_style_path(path)

  # Expect no warning is issued https://stackoverflow.com/questions/22003306/is-there-something-in-testthat-like-expect-no-warnings
  tab <- xltabr::initialise()
  testthat::expect_warning(xltabr:::style_catalogue_initialise(tab), regexp = NA)

  xltabr::set_style_path()
})
