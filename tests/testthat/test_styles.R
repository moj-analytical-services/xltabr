test_that("style conversion functions work as expected", {

  style_key_test <- "fontName_Calibri|fontSize_12|textDecoration_BOLD%ITALIC"

  t1 <- xltabr:::style_key_parser(style_key_test)

  expect_true(all(names(t1) == c("fontName", "fontSize", "textDecoration")))
  expect_true(t1[["fontName"]] == "Calibri")
  expect_true(t1[["fontSize"]] == "12")
  expect_true(all(t1[["textDecoration"]] == c("BOLD","ITALIC")))

  expect_true(xltabr:::create_style_key(t1) == style_key_test)

})
