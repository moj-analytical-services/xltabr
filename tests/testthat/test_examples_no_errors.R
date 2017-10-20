context("Test full examples don't produce unknown errors")

test_that("Test cross tab from synthetic data 1", {

  path <- system.file("extdata", "synthetic_data.csv", package="xltabr")
  df <- read.csv(path, stringsAsFactors = FALSE)
  ct <- reshape2::dcast(df, drive + age + colour ~ type, value.var= "value", margins=c("drive", "age", "colour"), fun.aggregate = mean)
  tab <- xltabr::auto_crosstab_to_wb(ct, return_tab = TRUE, fill_non_values_with = list(na = "**", nan = ".."))
})

test_that("Test cross tab from synthetic data 1.5 (adding NA and NaN)", {

  path <- system.file("extdata", "synthetic_data.csv", package="xltabr")
  df <- read.csv(path, stringsAsFactors = FALSE)
  ct <- reshape2::dcast(df, drive + age + colour ~ type, value.var= "value", margins=c("drive", "age", "colour"), fun.aggregate = mean)
  ct[4,4] <- NA
  ct[5,4] <- NaN
  ct[6,4] <- Inf
  ct[7,4] <- -Inf

  tab <- xltabr::auto_crosstab_to_wb(ct, return_tab = TRUE, fill_non_values_with = list(na = '*', nan = 10.1234, inf = '££', neg_inf = '$$'))
  expect_error(xltabr::auto_crosstab_to_wb(ct, return_tab = TRUE, fill_non_values_with = list(na = list('*',"**"))))
  expect_error(xltabr::auto_crosstab_to_wb(ct, return_tab = TRUE, fill_non_values_with = list(chicken = '*')))
})

test_that("Test cross tab from synthetic data 2", {

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
})

test_that("Test basic table", {
  tab <- xltabr::auto_df_to_wb(mtcars, return_tab=TRUE)

})

test_that("Test table numtypes", {
  path <- system.file("extdata", "test_number_types.rds", package="xltabr")
  df <- readRDS(path)
  tab <- xltabr::initialise()
  tab <- xltabr::add_body(tab, df)
  tab <- xltabr::auto_style_number_formatting(tab)
  tab <- xltabr:::write_all_elements_to_wb(tab)
  tab <- xltabr:::add_styles_to_wb(tab)

  tab <- xltabr::auto_df_to_wb(df, return_tab=TRUE)


})


