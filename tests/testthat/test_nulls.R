context("Test when null elements given")

test_that("Test elements not null", {

  tab <- xltabr::initialise()

  t1 = is.null(xltabr:::extent_get_rows(tab))
  t2 = is.null(xltabr:::extent_get_cols(tab))
  t3 = xltabr:::extent_get_bottom_wb_row(tab) == 0
  t4 = xltabr:::extent_get_rightmost_wb_col(tab) == 0

  testthat::expect_true(t1)
  testthat::expect_true(t2)
  testthat::expect_true(t3)
  testthat::expect_true(t4)

  t1 = nrow(xltabr:::title_get_cell_styles_table(tab)) == 0
  t2 = nrow(xltabr:::top_headers_get_cell_styles_table(tab)) == 0
  t3 = nrow(xltabr:::body_get_cell_styles_table(tab)) == 0
  t4 = nrow(xltabr:::footer_get_cell_styles_table(tab)) == 0
  t5 = nrow(xltabr:::combine_all_styles(tab)) == 0

  testthat::expect_true(t1)
  testthat::expect_true(t2)
  testthat::expect_true(t3)
  testthat::expect_true(t4)
  testthat::expect_true(t5)

  tab <- xltabr::initialise() %>%
         xltabr::add_footer("hello")

  t1 = xltabr:::extent_get_rows(tab) == 1
  t2 = xltabr:::extent_get_cols(tab) == 1
  t3 = xltabr:::extent_get_bottom_wb_row(tab) == 1
  t4 = xltabr:::extent_get_rightmost_wb_col(tab) == 1

  testthat::expect_true(t1)
  testthat::expect_true(t2)
  testthat::expect_true(t3)
  testthat::expect_true(t4)



})
