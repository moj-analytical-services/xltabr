context("Test autodetection")

test_that("Test meta columns are populated", {

  path <- system.file("extdata", "test_autodetect.csv", package="xltabr")

  df <- utils::read.csv(path, stringsAsFactors = FALSE)
  df$f <- as.Date(df$f)

  tab <- xltabr::initialise() %>%
    xltabr::add_body(df)

  tab <- xltabr:::auto_detect_left_headers(tab)

  t1 <- all(tab$body$left_header_colnames == c("a", "b", "c"))
  expect_true(t1)


  tab <- xltabr:::auto_detect_body_title_level(tab)

  t1 = all(tab$body$body_df$meta_row_ ==c("body|title_3", "body|title_4", "body|title_5", "body"))
  t2 = all(tab$body$body_df$meta_left_header_row_ == c("body|left_header|title_3", "body|left_header|title_4", "body|left_header|title_5",
                                                  "body|left_header"))
  expect_true(t1)
  expect_true(t2)

  tab <- auto_style_number_formatting(tab)

  t1 <- all(tab$body$meta_col_ == c("text1", "text1", "text1", "integer1",
                             "number1", "date1"))

  expect_true(t1)

  path <- system.file("extdata", "test_autodetect.csv", package="xltabr")

  df <- read.csv(path, stringsAsFactors = FALSE)

  tab <- xltabr::initialise() %>%
    xltabr::add_body(df) %>%
    xltabr:::auto_detect_left_headers() %>%
    xltabr:::auto_style_body_rows() %>%
    xltabr:::auto_style_indent() %>%
    xltabr::auto_style_number_formatting()


  xltabr:::body_get_cell_styles_table(tab)

})

test_that("Test that indent/coalesce works correctly", {
  path <- system.file("extdata", "test_autodetect.csv", package="xltabr")

  df <- read.csv(path, stringsAsFactors = FALSE)
  df$f <- as.Date(df$f)

  tab <- xltabr::initialise() %>%
    xltabr::add_body(df)

  tab <- xltabr:::auto_detect_left_headers(tab)

  t1 = all(tab$body$left_header_colnames == c("a", "b", "c"))
  testthat::expect_true(t1)

  tab <- xltabr:::auto_detect_body_title_level(tab)
  tab <- xltabr:::auto_style_indent(tab)

  #Check it's autodetected left_header_coluns correctly
  t1 = tab$body$left_header_colnames == tab$misc$coalesce_left_header_colname
  testthat::expect_true(t1)

  t1 = all(tab$body$body_df$meta_left_header_row_ ==  c("body|left_header|title_3", "body|left_header|title_4|indent_1",
                                              "body|left_header|title_5|indent_2", "body|left_header|indent_3"))
  testthat::expect_true(t1)

  t1 = all(tab$body$body_df_to_write$new_left_headers == c("Grand Total", "a", "b", "c"))

  testthat::expect_true(t1)



})


test_that("Test second autodetect dataset", {
  path <- system.file("extdata", "test_autodetect_2.csv", package="xltabr")

  df <- read.csv(path, stringsAsFactors = FALSE)
  df$g <- as.Date(df$g)

  cols <- colnames(df)
  tab <- xltabr::initialise() %>%
    xltabr::add_top_headers(cols) %>%
    xltabr::add_body(df)

  tab <- xltabr:::auto_detect_left_headers(tab)

  t1 <- all(tab$body$left_header_colnames == c("a", "b", "c", "d"))
  testthat::expect_true(t1)

  tab <- xltabr:::auto_detect_body_title_level(tab)
  tab <- xltabr:::auto_style_indent(tab)
  t1 = all(tab$body$body_df$meta_left_header_row_ == c("body|left_header|title_2", "body|left_header|title_3|indent_1",
                                              "body|left_header|title_4|indent_2", "body|left_header|title_5|indent_3",
                                              "body|left_header|indent_4"))

  testthat::expect_true(t1)

  t1 = (tab$body$left_header_colnames == " ")
  testthat::expect_true(t1)

  t1 = all(tab$top_headers$top_headers_list[[1]] == c(" ", "e", "f", "g"))
  testthat::expect_true(t1)

})
