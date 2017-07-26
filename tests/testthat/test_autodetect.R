context("Test body")

test_that("Test meta columns are populated", {

  path <- system.file("extdata", "test_autodetect.csv", package="xltabr")

  df <- readr::read_csv(path)

  tab <- xltabr::initialise() %>%
    xltabr::add_body(df)


  tab <- xltabr:::auto_detect_left_headers(tab)

  t1 <- all(tab$body$left_header_colnames == c("a", "b", "c"))
  expect_true(t1)


  tab <- xltabr:::auto_detect_body_title_level(tab)
  t1 = all(tab$body$body_df$meta_row_ ==c("body|title_3", "body|title_2", "body|title_1", "body"))
  t2 = all(tab$body$body_df$meta_left_header_row_ == c("body|left_header|title_3", "body|left_header|title_2", "body|left_header|title_1",
                                                  "body|left_header"))

  expect_true(t1)
  expect_true(t2)

  tab <- auto_style_number_formatting(tab)

  t1 <- all(tab$body$meta_col_==c("character1", "character1", "character1", "integer1",
                             "floating1", "date1"))

  expect_true(t1)


})
