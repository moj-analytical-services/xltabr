path <- system.file("extdata", "test_autodetect.csv", package="xltabr")

df <- readr::read_csv(path)

tab <- xltabr::initialise() %>%
  xltabr::add_body(df) %>%
  xltabr:::auto_detect_left_headers() %>%
  xltabr:::auto_detect_body_title_level() %>%
  xltabr:::auto_style_indent()



xltabr:::body_write_rows(tab)
xltabr:::body_get_cell_styles_table(tab)


tab <- xltabr:::add_styles_to_wb(tab, add_from = c("body"))

openxlsx::openXL(tab$wb)
