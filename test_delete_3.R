library(magrittr)
path <- system.file("extdata", "test_autodetect.csv", package="xltabr")

df <- read.csv(path, stringsAsFactors = FALSE)

# Leaving it here for now
# get_column_meta <- function(df){
#   meta_cols_ <- NULL
#   for (i in 1:length(df)){
#     meta_cols_ <- c(meta_cols_, ifelse(class(unlist(df[i])) == "character","character","number"))
#   }
#   return(meta_cols_)
# }

# mc <- get_column_meta(df)

tab <- xltabr::initialise() %>%
  xltabr::add_body(df) %>%
  xltabr:::auto_detect_left_headers() %>%
  xltabr:::auto_detect_body_title_level() %>%
  xltabr:::auto_style_indent() %>%
  xltabr:::auto_style_number_formatting()

tab$body$meta_col_

xltabr:::body_write_rows(tab)
xltabr:::body_get_cell_styles_table(tab)


tab <- xltabr:::add_styles_to_wb(tab, add_from = c("body"))

openxlsx::saveWorkbook(tab$wb, "testoutput.xlsx")

openxlsx::openXL(tab$wb)

ct <- reshape2::dcast(mtcars, am + gear ~ cyl, value.var= "mpg", margins=c("am", "gear"), fun.aggregate = mean)

wb  <- xltabr::auto_crosstab_to_wb(ct)

