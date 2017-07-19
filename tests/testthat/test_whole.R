context("Test whole")
library(magrittr)
library(testthat)

test_that("Simple test of dimensions, title", {

  title_text <- c("Title", "Subtitle", "")
  title_style_names <- c("title", "subtitle", "spacer")

  r1 <- paste0("r1_col_", as.character(1:2))
  r2 <- paste0("r2_col_", as.character(1:2))
  h_list <- list( r1,  r2)

  path <- system.file("extdata", "test_3x2.csv", package="xltabr")
  df <- readr::read_csv(path)

  # TODO:  Add footers to tests
  # footer_text <- c("Caveat 1 goes here", "Caveat 2 goes here")
  # footer_style_names <- c("footer", "footer")

  tab_all_elements_no_offset <- xltabr::initialise() %>%
    xltabr::add_title(title_text, title_style_names) %>%
    xltabr::add_top_headers(h_list) %>%
    xltabr::add_body(df, left_header_colnames = "a")

  tab_no_title <- xltabr::initialise() %>%
    xltabr::add_top_headers(h_list) %>%
    xltabr::add_body(df, left_header_colnames = "a")

  tab_all_elements_offset <- xltabr::initialise(topleft_row = 10, topleft_col  = 4) %>%
    xltabr::add_title(title_text, title_style_names) %>%
    xltabr::add_top_headers(h_list) %>%
    xltabr::add_body(df, left_header_colnames = "a")

  # Check titles reported correctly

  # Check bottom row is reported correctly
  t1 = xltabr:::title_get_bottom_wb_row(tab_all_elements_no_offset) == 3
  t2 = xltabr:::title_get_bottom_wb_row(tab_no_title) == 0
  t3 = xltabr:::title_get_bottom_wb_row(tab_all_elements_offset) == 12

  testthat::expect_true(t1)
  testthat::expect_true(t2)
  testthat::expect_true(t3)

  # Check the rightmost column is reported correctly
  t1 = xltabr:::title_get_rightmost_wb_col(tab_all_elements_no_offset) == 1
  t2 = xltabr:::title_get_rightmost_wb_col(tab_no_title) == 0
  t3 = xltabr:::title_get_rightmost_wb_col(tab_all_elements_offset) == 4

  testthat::expect_true(t1)
  testthat::expect_true(t2)
  testthat::expect_true(t3)

  # Check that the cols are reported correctly
  t1 = xltabr:::title_get_wb_cols(tab_all_elements_no_offset) == 1
  t2 = identical(xltabr:::title_get_wb_cols(tab_no_title), integer(0)) #Because you can't do integer(0) == integer(0)
  t3 = xltabr:::title_get_wb_cols(tab_all_elements_offset) == 4

  testthat::expect_true(t1)
  testthat::expect_true(t2)
  testthat::expect_true(t3)

  # Check that the rows are reported correctly
  t1 = xltabr:::title_get_wb_rows(tab_all_elements_no_offset) == 1:3
  t2 = identical(xltabr:::title_get_wb_rows(tab_no_title) ,integer(0))
  t3 = xltabr:::title_get_wb_rows(tab_all_elements_offset) == 10:12

  testthat::expect_true(all(t1))
  testthat::expect_true(t2)
  testthat::expect_true(all(t3))

  xltabr:::title_get_bottom_wb_row(tab_all_elements_offset)

  testthat::expect_true(all(tab_all_elements_no_offset$body$body_df$meta_row_ == c("body", "body", "body") ))
  testthat::expect_true(all(tab_all_elements_no_offset$body$body_df$meta_left_header_row_ == c("left_header_1", "left_header_1", "left_header_1") ))
  testthat::expect_true(all(tab_all_elements_no_offset$body$meta_col_ == c("", "") ))

  # Test styles derived correctly
  df2 <- xltabr:::body_get_cell_styles_table(tab_all_elements_no_offset)
  tab_all_elements_no_offset$body$body_df
  testthat::expect_true(all(df2$style_name == c("body|left_header_1", "body|left_header_1","body|left_header_1", "body", "body", "body")))


})

