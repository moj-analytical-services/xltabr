context("Test styles")

test_that("style conversion functions work as expected", {

  style_key_test <- "structure(list(fontName = structure(list(val = \"Calibri\"), .Names = \"val\"), fontSize = structure(list(val = 12), .Names = \"val\"), fontDecoration = c(\"BOLD\", \"ITALIC\")), .Names = c(\"fontName\", \"fontSize\", \"fontDecoration\"))"

  t1 <- xltabr:::style_key_parser(style_key_test)

  expect_true(all(names(t1) == c("fontName", "fontSize", "fontDecoration")))
  expect_true(t1[["fontName"]] == "Calibri")
  expect_true(t1[["fontSize"]] == 12)
  expect_true(all(t1[["fontDecoration"]] == c("BOLD","ITALIC")))

  expect_true(xltabr:::create_style_key(t1) == style_key_test)

})

test_that("Run through style_catalogue functions from reading style sheet to adding to final style_catalogue and wb", {

  path <- system.file("extdata", "tester_styles.xlsx", package = "xltabr")
  cell_path <- system.file("extdata", "style_to_excel_number_format_alt.csv", package = "xltabr")

  xltabr:::set_style_path(path)
  xltabr:::set_cell_format_path(cell_path)

  expected_style_names <- c("border1", "border2", "bg1", "text_colour1", "font1", "font2", "bg2", "text_colour2","number1","integer1","text1","date1","datetime1","general", "percent1")

  tab <- xltabr:::style_catalogue_initialise(list())

  expect_true(all(names(tab$style_catalogue) %in% expected_style_names))

  test_cell_style_def1 <- "border1|border2|integer1"
  expected_bs_1 <- "structure(list(fontName = structure(\"Calibri\", .Names = \"val\"), fontColour = structure(\"1\", .Names = \"theme\"), fontSize = structure(\"12\", .Names = \"val\"), fontFamily = structure(\"2\", .Names = \"val\"), fontScheme = structure(\"minor\", .Names = \"val\"), fontDecoration = c(\"BOLD\", \"ITALIC\"), borderRight = \"thin\", borderBottom = \"thin\", borderRightColour = structure(list( auto = \"1\"), .Names = \"auto\"), borderBottomColour = structure(list( auto = \"1\"), .Names = \"auto\"), borderTop = \"thin\", borderLeft = \"thin\", borderTopColour = structure(list(auto = \"1\"), .Names = \"auto\"), borderLeftColour = structure(list(auto = \"1\"), .Names = \"auto\"), numFmt = \"#,#;#,#;;* @\"), .Names = c(\"fontName\", \"fontColour\", \"fontSize\", \"fontFamily\", \"fontScheme\", \"fontDecoration\", \"borderRight\", \"borderBottom\", \"borderRightColour\", \"borderBottomColour\", \"borderTop\", \"borderLeft\", \"borderTopColour\", \"borderLeftColour\", \"numFmt\"))"

  test_cell_style_def2 <- "border1|bg1|text_colour1|font1|text1"
  expected_bs_2 <- "structure(list(fontName = structure(\"Arial\", .Names = \"val\"), fontColour = structure(\"1\", .Names = \"theme\"), fontSize = structure(\"16\", .Names = \"val\"), fontFamily = structure(\"2\", .Names = \"val\"), fontScheme = structure(\"minor\", .Names = \"val\"), fontDecoration = \"BOLD\", borderRight = \"thin\", borderBottom = \"thin\", borderRightColour = structure(list(auto = \"1\"), .Names = \"auto\"), borderBottomColour = structure(list(auto = \"1\"), .Names = \"auto\"), fill = structure(list(fillFg = structure(\"FFFFFF00\", .Names = \" rgb\"), fillBg = structure(\"64\", .Names = \" indexed\")), .Names = c(\"fillFg\", \"fillBg\")), numFmt = \"TEXT\"), .Names = c(\"fontName\", \"fontColour\", \"fontSize\", \"fontFamily\", \"fontScheme\", \"fontDecoration\", \"borderRight\", \"borderBottom\", \"borderRightColour\", \"borderBottomColour\", \"fill\", \"numFmt\"))"

  test_cell_style_def3 <- "border2|bg2|text_colour2|font2|date1"
  expected_bs_3 <- "structure(list(fontName = structure(\"Times New Roman\", .Names = \"val\"), fontColour = structure(\"1\", .Names = \"theme\"), fontSize = structure(\"14\", .Names = \"val\"), fontScheme = structure(\"minor\", .Names = \"val\"), fontDecoration = \"ITALIC\", borderTop = \"thin\", borderLeft = \"thin\", borderTopColour = structure(list( auto = \"1\"), .Names = \"auto\"), borderLeftColour = structure(list( auto = \"1\"), .Names = \"auto\"), fill = structure(list( fillFg = structure(\"FF00FFE4\", .Names = \" rgb\"), fillBg = structure(\"64\", .Names = \" indexed\")), .Names = c(\"fillFg\", \"fillBg\")), numFmt = \"yyyy-mm-dd\"), .Names = c(\"fontName\", \"fontColour\", \"fontSize\", \"fontScheme\", \"fontDecoration\", \"borderTop\", \"borderLeft\", \"borderTopColour\", \"borderLeftColour\", \"fill\", \"numFmt\"))"

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
