library(magrittr)
test_that("test row heights work", {

  path <- system.file("extdata", "styles_rowheight.xlsx", package = "xltabr")

  xltabr::set_style_path(path)

  df <- data.frame("a" = 1:6, "b" = 4:9)

  tab <- initialise() %>%
    add_body(df)

  tab$body$body_df$meta_row_ <- c("height_100", "height_45", "height_blank", "height_100", "height_blank|height_45", "height_45|height_100")

  xltabr:::write_all_elements_to_wb(tab)

  xltabr:::add_styles_to_wb(tab)

  expected <- c("100", "45", "100", "45", "100")
  names(expected) <- c(1, 2, 4, 5, 6)


  t1 <- tab$wb$rowHeights[[1]] == expected

  testthat::expect_true(all(t1))

})



test_that("test row heights work 2", {

  path <- system.file("extdata", "styles_rowheight.xlsx", package = "xltabr")

  xltabr::set_style_path(path)

  df <- data.frame("a" = 1:3, "b" = 4:6, "c" = 14:16)

  tab <- initialise() %>%
    add_body(df)

  tab$body$body_df$meta_row_ <- c("height_100", "height_45", "height_blank")
  tab$body$meta_col_ <- c("height_blank", "height_45", "height_45")

  xltabr:::write_all_elements_to_wb(tab)

  xltabr:::add_styles_to_wb(tab)

  expected <- c("100", "45", "45")
  names(expected) <- c(1, 2, 3)


  t1 <- tab$wb$rowHeights[[1]] == expected

  testthat::expect_true(all(t1))

})
