context("Test body")

test_that("Test meta columns are populated", {

  path <- system.file("extdata", "test_2x2.csv", package="xltabr")

  df <- read.csv(path, stringsAsFactors = FALSE)

  tab <- xltabr::initialise() %>%
    xltabr::add_body(df, left_header_colnames = "a")

  testthat::expect_true(all(tab$body$body_df$meta_row_ == c("body", "body") ))
  testthat::expect_true(all(tab$body$body_df$meta_left_header_row_ == c("body|left_header", "body|left_header") ))
  testthat::expect_true(all(tab$body$meta_col_ == c("", "") ))

  # Test styles derived correctly
  df2 <- xltabr:::body_get_cell_styles_table(tab)
  testthat::expect_true(all(df2$style_name == c("body|left_header", "body|left_header", "body", "body")))

})



test_that("Test user addition of style information", {

  path <- system.file("extdata", "test_3x3.csv", package="xltabr")
  df <- read.csv(path, stringsAsFactors = FALSE)

  # Adding a single style should work
  tab <- xltabr::initialise() %>%
    xltabr::add_body(df, left_header_colnames = "a" , row_style_names = "b", left_header_style_names = "lh")

  t1 <- all(xltabr:::body_get_cell_styles_table(tab)$style_name == c(rep("lh",3), rep("b", 6)))
  testthat::expect_true(t1)

  # Adding two should raise warning; it does not object multiplication rules
  tab <- xltabr::initialise()
  testthat::expect_warning(xltabr::add_body(tab, df, left_header_colnames = "a" , row_style_names = c("br", "br2"), left_header_style_names = c("lhr", "lhr2")))

  # Adding three is the core use case and should work
  rsn <- c("br1", "br2", "br3")
  lhsn <-c("lh1", "lh2", "lh3")

  tab <- xltabr::initialise() %>%
    xltabr::add_body(df, left_header_colnames = "a" , row_style_names = rsn, left_header_style_names = lhsn)

  t1 <- all(xltabr:::body_get_cell_styles_table(tab)$style_name  == c(lhsn, rsn, rsn))
  testthat::expect_true(t1)

  # Similarly, with col_style_names
  csn <- "col"
  tab <- xltabr::initialise() %>%
    xltabr::add_body(df, left_header_colnames = "a" , row_style_names = rsn, left_header_style_names = lhsn, col_style_names = csn)

  style_names <- paste(c(lhsn, rsn, rsn), "col", sep="|")

  t1 = all(xltabr:::body_get_cell_styles_table(tab)$style_name  == style_names)
  testthat::expect_true(t1)

  # Similarly, with col_style_names
  csn <- c("col", "col1")
  tab <- xltabr::initialise()
  testthat::expect_warning(xltabr::add_body(tab, df, left_header_colnames = "a" , row_style_names = rsn, left_header_style_names = lhsn, col_style_names = csn))

  #Provide all three col style names
  csn <- c("col1", "col2", "col3")
  tab <- xltabr::initialise() %>%
    xltabr::add_body(df, left_header_colnames = "a" , row_style_names = rsn, left_header_style_names = lhsn, col_style_names = csn)

  style_names <- paste(c(lhsn, rsn, rsn), rep(csn, each = 3), sep = "|")

  t1 = all(xltabr:::body_get_cell_styles_table(tab)$style_name  == style_names)

  testthat::expect_true(t1)


  })






