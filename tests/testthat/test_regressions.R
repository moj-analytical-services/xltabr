# This script contains regression tests corresponding to issues in github

# https://github.com/moj-analytical-services/xltabr/issues/31
context("Test regressions")
library(magrittr)

test_that("test styles table does not have factors", {
  path <- system.file("extdata", "synthetic_data.csv", package="xltabr")
  df <- read.csv(path, stringsAsFactors = FALSE)
  ct <- reshape2::dcast(df, drive + age + colour ~ type, value.var= "value", margins=c("drive", "age", "colour"), fun.aggregate = mean)
  tab <- xltabr::auto_crosstab_to_wb(ct, return_tab = TRUE, titles = "title", footers = "footers")

  df <- xltabr:::combine_all_styles(tab)
  t1 = (class(df$style_name) == "character")
  testthat::expect_true(t1)
})
