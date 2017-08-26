context("Test styles")

test_that("style conversion functions work as expected", {

  style_key_test <- "fontName_Calibri|fontSize_12|textDecoration_BOLD%ITALIC"

  t1 <- xltabr:::style_key_parser(style_key_test)

  expect_true(all(names(t1) == c("fontName", "fontSize", "textDecoration")))
  expect_true(t1[["fontName"]] == "Calibri")
  expect_true(t1[["fontSize"]] == "12")
  expect_true(all(t1[["textDecoration"]] == c("BOLD","ITALIC")))

  expect_true(xltabr:::create_style_key(t1) == style_key_test)

})

test_that("Run through style_catalogue functions from reading style sheet to adding to final style_catalogue and wb", {

  path <- system.file("extdata", "tester_styles.xlsx", package = "xltabr")
  num_path <- system.file("extdata", "style_to_excel_number_format_alt.csv", package = "xltabr")

  expected_style_names <- c("border1", "border2", "bg1", "text_colour1", "font1", "font2", "bg2", "text_colour2","number1","integer1","text1","date1","datetime1","general")

  tab <- xltabr:::style_catalogue_initialise(list(), styles_xlsx = path, num_styles_csv = num_path)

  expect_true(all(names(tab$style_catalogue) %in% expected_style_names))

  test_cell_style_def1 <- "border1|border2|integer1"
  expected_bs_1 <- "fontName_Calibri|fontSize_12|fontColour_1|numFmt_#,#|borderBottom_thin|borderBottomColour_1|borderTop_thin|borderTopColour_1|borderLeft_thin|borderLeftColour_1|borderRight_thin|borderRightColour_1|textDecoration_BOLD%ITALIC"

  test_cell_style_def2 <- "border1|bg1|text_colour1|font1|text1"
  expected_bs_2 <- "fontName_Arial|fontSize_16|fontColour_1|numFmt_TEXT|borderBottom_thin|borderBottomColour_1|borderRight_thin|borderRightColour_1|bgFill_64|fgFill_FFFFFF00|textDecoration_BOLD"

  test_cell_style_def3 <- "border2|bg2|text_colour2|font2|date1"
  expected_bs_3 <- "fontName_Times New Roman|fontSize_14|fontColour_1|numFmt_yyyy-mm-dd|borderTop_thin|borderTopColour_1|borderLeft_thin|borderLeftColour_1|bgFill_64|fgFill_FF00FFE4|textDecoration_ITALIC"

  bs1 <- xltabr:::build_style(tab, cell_style_definition = test_cell_style_def1)
  bs2 <- xltabr:::build_style(tab, cell_style_definition = test_cell_style_def2)
  bs3 <- xltabr:::build_style(tab, cell_style_definition = test_cell_style_def3)

  expect_true(expected_bs_1 == bs1)
  expect_true(expected_bs_2 == bs2)
  expect_true(expected_bs_3 == bs3)

  multiple_styles <- c(test_cell_style_def1, test_cell_style_def2, test_cell_style_def3)
  tab <- xltabr:::add_style_defintions_to_catelogue(tab, multiple_styles)

  expect_true(all(names(tab$style_catalogue) %in% c(expected_style_names, multiple_styles)))
})

