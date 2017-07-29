context("Test body")

test_that("Test meta columns are populated", {

  path <- system.file("extdata", "test_2x2.csv", package="xltabr")

  df <- readr::read_csv(path)

  tab <- xltabr::initialise() %>%
    xltabr::add_body(df, left_header_colnames = "a")

  testthat::expect_true(all(tab$body$body_df$meta_row_ == c("body", "body") ))
  testthat::expect_true(all(tab$body$body_df$meta_left_header_row_ == c("body|left_header", "body|left_header") ))
  testthat::expect_true(all(tab$body$meta_col_ == c("", "") ))

  # Test styles derived correctly
  df2 <- xltabr:::body_get_cell_styles_table(tab)
  testthat::expect_true(all(df2$style_name == c("body|left_header", "body|left_header", "body", "body")))




})

test_that("Test dimentions for super simple 2x2", {

  path <- system.file("extdata", "test_2x2.csv", package="xltabr")
  df <- readr::read_csv(path)

  tab <- xltabr::initialise() %>%
    xltabr::add_body(df)

  cols <- xltabr:::body_get_wb_cols(tab)
  testthat::expect_true(all(cols == 1:2))

  testthat::expect_true(all(xltabr:::body_get_wb_rows(tab) == 1:2))
  testthat::expect_true(xltabr:::body_get_bottom_wb_row(tab) == 2)
  testthat::expect_true(xltabr:::body_get_rightmost_wb_col(tab) == 2)
  xltabr:::body_get_rightmost_wb_col(tab)


})



