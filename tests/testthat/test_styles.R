context("Test styles")

test_that("Test font decoration inherits correctly", {

  path <- system.file("extdata", "tester_styles.xlsx", package = "xltabr")
  xltabr::set_style_path(path)
  tab <- xltabr::initialise()

  tab <- xltabr:::create_style_if_not_exists(tab, "bold|italic")
  xltabr_style <- tab$style_catalogue$`bold|italic`

  # Check bolditalic works as anticipated etc
  t1 <- all(xltabr_style$style_list$fontDecoration == c("BOLD", "ITALIC"))
  expect_true(t1)

  t2 <- all(xltabr_style$s4style$fontDecoration == c("BOLD", "ITALIC"))
  expect_true(t2)

  # Manual test that it actually renders to bold_italic - commented so test does not open Excel!
  # wb <- openxlsx::createWorkbook()
  # openxlsx::addWorksheet(wb, "Cars")
  # openxlsx::writeData(wb, "Cars", "hello")
  # openxlsx::addStyle(wb, "Cars", xltabr_style$s4style, 1,1)
  # openxlsx::openXL(wb)

})

test_that("Check row height inherits as expected", {

  path <- system.file("extdata", "tester_styles.xlsx", package = "xltabr")
  xltabr::set_style_path(path)
  tab <- xltabr::initialise()

  # Check we take the max height
  style_string <- "border1|rowheight_30|rowheight_40"
  tab <- xltabr:::create_style_if_not_exists(tab, style_string)
  xltabr_style <- tab$style_catalogue[[style_string]]
  t1 <- xltabr_style$row_height == 40
  expect_true(t1)

  # Order shouldn't matter
  style_string <- "border1|rowheight_40|font1|rowheight_30|font2"
  tab <- xltabr:::create_style_if_not_exists(tab, style_string)
  xltabr_style <- tab$style_catalogue[[style_string]]
  t2 <- xltabr_style$row_height == 40
  expect_true(t2)

  # If no height specified in any, should return null
  style_string <- "bolditalic|bg1"
  tab <- xltabr:::create_style_if_not_exists(tab, style_string)
  xltabr_style <- tab$style_catalogue[[style_string]]

  t3 <- is.null(xltabr_style$row_height)
  expect_true(t3)

  # Manual test that it actually renders to yellow - commented so test does not open Excel!
  # tab <- xltabr::initialise()
  # wb <- tab$wb
  # openxlsx::writeData(wb, "Sheet1", "hello")
  # style_string <- "`border1|rowheight_40|font1|rowheight_30|font2"
  # tab <- xltabr:::create_style_if_not_exists(tab, style_string)
  # xltabr_style <- tab$style_catalogue[[style_string]]
  # openxlsx::addStyle(wb, "Sheet1", xltabr_style$s4style, 1,1)
  # tab <- xltabr:::update_row_heights(tab, 1, xltabr_style)
  # openxlsx::openXL(tab$wb)


})

test_that("Check background colour works", {

  path <- system.file("extdata", "tester_styles.xlsx", package = "xltabr")
  xltabr::set_style_path(path)
  tab <- xltabr::initialise()

  style_string <- "bg1|bg2"
  tab <- xltabr:::create_style_if_not_exists(tab, style_string)
  xltabr_style <- tab$style_catalogue[[style_string]]
  t1 <- xltabr_style$style_list$fill$fillFg == "FF00FFE4"
  expect_true(t1)

  style_string <- "bg2|bg1|border1"
  tab <- xltabr:::create_style_if_not_exists(tab, style_string)
  xltabr_style <- tab$style_catalogue[[style_string]]
  t2 <- xltabr_style$style_list$fill$fillFg == "FFFFFF00"
  expect_true(t2)

  # Manual test that it actually renders to yellow - commented so test does not open Excel!
  # wb <- openxlsx::createWorkbook()
  # openxlsx::addWorksheet(wb, "Cars")
  # openxlsx::writeData(wb, "Cars", "hello")
  # openxlsx::addStyle(wb, "Cars", xltabr_style$s4style, 1,1)
  # openxlsx::openXL(wb)

})


test_that("Test that border inheritence works" ,{

  path <- system.file("extdata", "tester_styles.xlsx", package = "xltabr")
  xltabr::set_style_path(path)
  tab <- xltabr::initialise()

  style_string <- "border1|border2|bold|font1|font2"
  tab <- xltabr:::create_style_if_not_exists(tab, style_string)
  xltabr_style <- tab$style_catalogue[[style_string]]

  expect_true(xltabr_style$style_list$borderBottom == "medium")
  expect_true(xltabr_style$style_list$borderRight == "thin")
  expect_true(xltabr_style$style_list$borderLeft == "thin")
  expect_true(xltabr_style$style_list$borderTop == "thin")


  # Manual test that it actually renders to bold_italic - commented so test does not open Excel!
  # wb <- openxlsx::createWorkbook()
  # openxlsx::addWorksheet(wb, "Cars")
  # openxlsx::writeData(wb, "Cars", "hello", 2,2)
  # openxlsx::addStyle(wb, "Cars", xltabr_style$s4style, 2,2)
  # openxlsx::openXL(wb)

})

test_that("Test that text colour works", {

  path <- system.file("extdata", "tester_styles.xlsx", package = "xltabr")
  xltabr::set_style_path(path)
  tab <- xltabr::initialise()

  style_string <- "text_colour1|text_colour2|bold|font1|font2"
  tab <- xltabr:::create_style_if_not_exists(tab, style_string)
  xltabr_style <- tab$style_catalogue[[style_string]]

  # Manual test that it actually renders to bold_italic - commented so test does not open Excel!
  # wb <- openxlsx::createWorkbook()
  # openxlsx::addWorksheet(wb, "Cars")
  # openxlsx::writeData(wb, "Cars", "hello", 2,2)
  # openxlsx::addStyle(wb, "Cars", xltabr_style$s4style, 2,2)
  # openxlsx::openXL(wb)

})



test_that("Check row height works across multiple sheets", {

  path <- system.file("extdata", "tester_styles.xlsx", package = "xltabr")
  xltabr::set_style_path(path)
  tab <- xltabr::initialise()

  path <- system.file("extdata", "test_3x3.csv", package = "xltabr")
  df <- read.csv(path)

  tab <- xltabr::add_body(tab, df)
  tab$body$body_df$meta_row_ <- c("italic", "border2|rowheight_30|rowheight_40", "font1|rowheight_30")
  tab$body$meta_col_ <- c("text_colour1", "bg2", "italic")
  tab <- xltabr:::write_all_elements_to_wb(tab)
  tab <- xltabr:::add_styles_to_wb(tab)

  tab <- xltabr::initialise(insert_below_tab = tab)

  tab <- xltabr::add_body(tab, df)
  tab$body$body_df$meta_row_ <- c("rowheight_40", "bg2", "bg1")
  tab$body$meta_col_ <- c("", "", "")
  tab <- xltabr:::write_all_elements_to_wb(tab)
  tab <- xltabr:::add_styles_to_wb(tab)

  tab <- xltabr::initialise(tab$wb, ws_name = "My_second_sheet")

  tab <- xltabr::add_body(tab, df)
  tab$body$body_df$meta_row_ <- c("rowheight_40", "rowheight_40", "bg2")
  tab$body$meta_col_ <- c("", "", "")
  tab <- xltabr:::write_all_elements_to_wb(tab)
  tab <- xltabr:::add_styles_to_wb(tab)

  #We now expect the row height on the second sheet of the workbook to be 40, 40, null

  t1 <- all(tab$wb$rowHeights[[1]] == c("40","30","40"))
  testthat::expect_true(t1)

  t1 <- all(tab$wb$rowHeights[[2]] == c("40", "40"))
  testthat::expect_true(t1)

  # Check this works in Excel
  # openxlsx::openXL(tab$wb)


})


test_that("Check number formats inherit as expected", {

  path <- system.file("extdata", "tester_styles.xlsx", package = "xltabr")
  xltabr::set_style_path(path)
  tab <- xltabr::initialise()

  style_string <- "percent1|integer1"
  tab <- xltabr:::create_style_if_not_exists(tab, style_string)
  xltabr_style <- tab$style_catalogue[[style_string]]
  df <- read.csv(xltabr::get_cell_format_path())
  expected <- as.character(df[df$style_name == "integer1",]$excel_format)
  t1 <- (xltabr_style$s4style$numFmt$formatCode == expected)

  style_string <- "integer1|percent1"
  tab <- xltabr:::create_style_if_not_exists(tab, style_string)
  xltabr_style <- tab$style_catalogue[[style_string]]
  df <- read.csv(xltabr::get_cell_format_path())
  expected <- as.character(df[df$style_name == "percent1",]$excel_format)
  t1 <- (xltabr_style$s4style$numFmt$formatCode == expected)

  # Manual test that it actually renders to bold_italic - commented so test does not open Excel!
  # wb <- openxlsx::createWorkbook()
  # openxlsx::addWorksheet(wb, "Cars")
  # openxlsx::writeData(wb, "Cars", 1.2, 2,2)
  # openxlsx::addStyle(wb, "Cars", xltabr_style$s4style, 2,2)
  # openxlsx::openXL(wb)

})


test_that("Check it's possible to add a number format", {

  path <- system.file("extdata", "tester_styles.xlsx", package = "xltabr")
  xltabr::set_style_path(path)
  tab <- xltabr::initialise()
  tab <- style_catalogue_add_excel_num_format(tab, "currency2", "£ #,###")

  t1 <- tab$style_catalogue$currency2$s4style$numFmt$numFmtId == 9999
  t2 <- tab$style_catalogue$currency2$s4style$numFmt$formatCode == "£ #,###"

  expect_true(t1)
  expect_true(t2)

  # Manual test that it actually renders to bold_italic - commented so test does not open Excel!
  # wb <- openxlsx::createWorkbook()
  # openxlsx::addWorksheet(wb, "Cars")
  # openxlsx::writeData(wb, "Cars", 12.12)
  # openxlsx::addStyle(wb, "Cars", tab$style_catalogue$currency2$s4style, 1,1)
  # openxlsx::openXL(wb)


})

test_that("Check it's possible to add a custom s4 style", {

  path <- system.file("extdata", "tester_styles.xlsx", package = "xltabr")
  xltabr::set_style_path(path)
  tab <- xltabr::initialise()

  s4style <- openxlsx::createStyle(fontName = "Courier", fontColour = "#80a9ed", fontSize = 20, numFmt =  "£    #,###")

  tab <- style_catalogue_add_openxlsx_style(tab, "custom", s4style, row_height = 40)

  t1 <- tab$style_catalogue$custom$row_height == 40
  t2 <- tab$style_catalogue$custom$style_list$fontName == "Courier"
  t3 <- tab$style_catalogue$custom$s4style$fontColour == "FF80A9ED"

  expect_true(t1)
  expect_true(t2)
  expect_true(t3)


  # Manual test that it actually renders to bold_italic - commented so test does not open Excel!
  # wb <- openxlsx::createWorkbook()
  # openxlsx::addWorksheet(wb, "Cars")
  # openxlsx::writeData(wb, "Cars", 12.12)
  # openxlsx::addStyle(wb, "Cars", tab$style_catalogue$custom$s4style, 1,1)
  # openxlsx::openXL(wb)



})


test_that("Have a look at what the various default number formats look like", {

  tab <- xltabr::initialise()
  xltabr::set_style_path()

  path <- get_cell_format_path()
  df <- read.csv(path, stringsAsFactors = FALSE)
  style_strings <- df$style_name

  # Create dataframe with numbers of different sizes
  nums <- 2.1 ** (1:20)

  l <- list()
  for (s in style_strings) {
    l[[s]] <- nums
  }

  df <- data.frame(l)

  tab <- auto_crosstab_to_tab(df)
  tab$body$meta_col_ <- style_strings

  tab <- xltabr::write_data_and_styles_to_wb(tab)

  # openxlsx::openXL(tab$wb)

})
