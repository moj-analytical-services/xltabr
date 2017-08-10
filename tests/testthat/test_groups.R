context("Test groups of elements (title, body etc.)")

suppressMessages( library(magrittr))
library(testthat)

title_text <- c("Title", "Subtitle", "")
title_style_names <- c("title", "subtitle", "spacer")

r1 <- paste0("r1_col_", as.character(1:2))
r2 <- paste0("r2_col_", as.character(1:2))
h_list <- list( r1,  r2)

path <- system.file("extdata", "test_3x2.csv", package="xltabr")
df <- read.csv(path, stringsAsFactors = FALSE)

footer_text <- c("Caveat 1 goes here", "Caveat 2 goes here")
footer_style_names <- c("footer", "footer")

tab_all_elements_no_offset <- xltabr::initialise() %>%
  xltabr::add_title(title_text, title_style_names) %>%
  xltabr::add_top_headers(h_list) %>%
  xltabr::add_body(df, left_header_colnames = "a") %>%
  xltabr::add_footer(footer_text, footer_style_names)

tab_no_title <- xltabr::initialise() %>%
  xltabr::add_top_headers(h_list) %>%
  xltabr::add_body(df, left_header_colnames = "a") %>%
  xltabr::add_footer(footer_text, footer_style_names)

tab_all_elements_offset <- xltabr::initialise(topleft_row = 10, topleft_col  = 4) %>%
  xltabr::add_title(title_text, title_style_names) %>%
  xltabr::add_top_headers(h_list) %>%
  xltabr::add_body(df, left_header_colnames = "a") %>%
  xltabr::add_footer(footer_text, footer_style_names)

tab_no_top_headers <- xltabr::initialise() %>%
  xltabr::add_title(title_text, title_style_names) %>%
  xltabr::add_body(df, left_header_colnames = "a") %>%
  xltabr::add_footer(footer_text, footer_style_names)


test_that("Simple test of dimensions, title", {

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
  t2 = is.null(xltabr:::title_get_wb_cols(tab_no_title))
  t3 = xltabr:::title_get_wb_cols(tab_all_elements_offset) == 4

  testthat::expect_true(t1)
  testthat::expect_true(t2)
  testthat::expect_true(t3)

  # Check that the rows are reported correctly
  t1 = xltabr:::title_get_wb_rows(tab_all_elements_no_offset) == 1:3
  t2 = is.null(xltabr:::title_get_wb_rows(tab_no_title) )
  t3 = xltabr:::title_get_wb_rows(tab_all_elements_offset) == 10:12

  testthat::expect_true(all(t1))
  testthat::expect_true(t2)
  testthat::expect_true(all(t3))

})



test_that("Simple test of dimensions, top_header", {

  # Check bottom row is reported correctly
  t1 = xltabr:::top_headers_get_bottom_wb_row(tab_all_elements_no_offset) == 5
  t2 = xltabr:::top_headers_get_bottom_wb_row(tab_no_title) == 2
  t3 = xltabr:::top_headers_get_bottom_wb_row(tab_all_elements_offset) == 14

  testthat::expect_true(t1)
  testthat::expect_true(t2)
  testthat::expect_true(t3)

  # Check the rightmost column is reported correctly
  t1 = xltabr:::top_headers_get_rightmost_wb_col(tab_all_elements_no_offset) == 2
  t2 = xltabr:::top_headers_get_rightmost_wb_col(tab_no_top_headers) == 0
  t3 = xltabr:::top_headers_get_rightmost_wb_col(tab_all_elements_offset) == 5

  testthat::expect_true(t1)
  testthat::expect_true(t2)
  testthat::expect_true(t3)

  # Check that the cols are reported correctly
  t1 = all(xltabr:::top_headers_get_wb_cols(tab_all_elements_no_offset) == 1:2)
  t2 = is.null(xltabr:::top_headers_get_wb_cols(tab_no_top_headers))
  t3 = all(xltabr:::top_headers_get_wb_cols(tab_all_elements_offset) == 4:5)

  testthat::expect_true(t1)
  testthat::expect_true(t2)
  testthat::expect_true(t3)

  # Check that the rows are reported correctly
  t1 = xltabr:::top_headers_get_wb_rows(tab_all_elements_no_offset) == 4:5
  t2 = is.null(xltabr:::top_headers_get_wb_rows(tab_no_top_headers))
  t3 = xltabr:::top_headers_get_wb_rows(tab_all_elements_offset) == 13:14

  testthat::expect_true(all(t1))
  testthat::expect_true(t2)
  testthat::expect_true(all(t3))


})

test_that("Simple test of dimensions, body", {

  # Check bottom row is reported correctly
  t1 = xltabr:::body_get_bottom_wb_row(tab_all_elements_no_offset) == 8
  t2 = xltabr:::body_get_bottom_wb_row(tab_no_title) == 5
  t3 = xltabr:::body_get_bottom_wb_row(tab_all_elements_offset) == 17

  testthat::expect_true(t1)
  testthat::expect_true(t2)
  testthat::expect_true(t3)

  # Check the rightmost column is reported correctly
  t1 = xltabr:::body_get_rightmost_wb_col(tab_all_elements_no_offset) == 2
  t2 = xltabr:::body_get_rightmost_wb_col(tab_no_top_headers) == 2
  t3 = xltabr:::body_get_rightmost_wb_col(tab_all_elements_offset) == 5

  testthat::expect_true(t1)
  testthat::expect_true(t2)
  testthat::expect_true(t3)

  # Check that the cols are reported correctly
  t1 = all(xltabr:::body_get_wb_cols(tab_all_elements_no_offset) == 1:2)

  t2 = all(xltabr:::body_get_wb_cols(tab_no_top_headers) == 1:2)
  t3 = all(xltabr:::body_get_wb_cols(tab_all_elements_offset) == 4:5)

  testthat::expect_true(t1)
  testthat::expect_true(t2)
  testthat::expect_true(t3)

  # Check that the rows are reported correctly
  t1 = xltabr:::body_get_wb_rows(tab_all_elements_no_offset) == 6:8
  t2 = all(xltabr:::body_get_wb_rows(tab_no_top_headers) == 4:6)
  t3 = xltabr:::body_get_wb_rows(tab_all_elements_offset) == 15:17

  testthat::expect_true(all(t1))
  testthat::expect_true(t2)
  testthat::expect_true(all(t3))

})


test_that("Simple test of dimensions, footer", {

  # Check bottom row is reported correctly
  t1 = xltabr:::footer_get_bottom_wb_row(tab_all_elements_no_offset) == 10
  t2 = xltabr:::footer_get_bottom_wb_row(tab_no_title) == 7
  t3 = xltabr:::footer_get_bottom_wb_row(tab_all_elements_offset) == 19

  testthat::expect_true(t1)
  testthat::expect_true(t2)
  testthat::expect_true(t3)

  # Check the rightmost column is reported correctly
  t1 = xltabr:::footer_get_rightmost_wb_col(tab_all_elements_no_offset) == 1
  t2 = xltabr:::footer_get_rightmost_wb_col(tab_no_top_headers) == 1
  t3 = xltabr:::footer_get_rightmost_wb_col(tab_all_elements_offset) == 4

  testthat::expect_true(t1)
  testthat::expect_true(t2)
  testthat::expect_true(t3)

  # Check that the cols are reported correctly
  t1 = all(xltabr:::footer_get_wb_cols(tab_all_elements_no_offset) == 1)
  t2 = all(xltabr:::footer_get_wb_cols(tab_no_top_headers) == 1)
  t3 = all(xltabr:::footer_get_wb_cols(tab_all_elements_offset) == 4)

  testthat::expect_true(t1)
  testthat::expect_true(t2)
  testthat::expect_true(t3)

  # Check that the rows are reported correctly
  t1 = all(xltabr:::footer_get_wb_rows(tab_all_elements_no_offset) == 9:10)
  t2 = all(xltabr:::footer_get_wb_rows(tab_no_top_headers) == 7:8)
  t3 = xltabr:::footer_get_wb_rows(tab_all_elements_offset) == 18:19

  testthat::expect_true(all(t1))
  testthat::expect_true(t2)
  testthat::expect_true(all(t3))

})

test_that("Test of combined dimension", {

  t1 = all(xltabr:::extent_get_rows(tab_all_elements_no_offset) == 1:10)
  t2 = all(xltabr:::extent_get_rows(tab_no_top_headers) == 1:8)
  t3 = all(xltabr:::extent_get_rows(tab_all_elements_offset) == 10:19)

  testthat::expect_true(t1)
  testthat::expect_true(t2)
  testthat::expect_true(t3)

  t1 = all(xltabr:::extent_get_cols(tab_all_elements_no_offset) == 1:2)
  t2 = all(xltabr:::extent_get_cols(tab_no_top_headers) == 1:2)
  t3 = all(xltabr:::extent_get_cols(tab_all_elements_offset) == 4:5)

  testthat::expect_true(t1)
  testthat::expect_true(t2)
  testthat::expect_true(t3)


  t1 = all(xltabr:::extent_get_bottom_wb_row(tab_all_elements_no_offset) == 10)
  t2 = all(xltabr:::extent_get_bottom_wb_row(tab_no_top_headers) == 8)
  t3 = all(xltabr:::extent_get_bottom_wb_row(tab_all_elements_offset) == 19)

  testthat::expect_true(t1)
  testthat::expect_true(t2)
  testthat::expect_true(t3)


  t1 = all(xltabr:::extent_get_rightmost_wb_col(tab_all_elements_no_offset) == 2)
  t2 = all(xltabr:::extent_get_rightmost_wb_col(tab_no_top_headers) == 2)
  t3 = all(xltabr:::extent_get_rightmost_wb_col(tab_all_elements_offset) == 5)

  testthat::expect_true(t1)
  testthat::expect_true(t2)
  testthat::expect_true(t3)


})

