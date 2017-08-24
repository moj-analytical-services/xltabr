library(xltabr)
library(dplyr)

# Read data from main database and run cross tabulation
path <- system.file("extdata", "synthetic_data.csv", package="xltabr")
df <- read.csv(path, stringsAsFactors = FALSE)
ct <- reshape2::dcast(df, drive + age + colour + type ~ date, value.var= "value", margins=c("drive", "age", "colour", "type"), fun.aggregate = sum)
ct <- ct %>% dplyr::arrange(-row_number())


# Example 1: default settings
wb <- xltabr::auto_crosstab_to_wb(ct, titles = "This is the title")
openxlsx::openXL(wb)

# Example 2:  User provides their own formatting options

path <- system.file("extdata", "styles_pub.xlsx", package = "xltabr")
num_path <- system.file("extdata", "style_to_excel_number_format_alt.csv", package = "xltabr")
title <- "This is the title"
footers <- c("Footer information 1", "Footer information 2")
tab <- xltabr::auto_crosstab_to_wb(ct, return_tab = TRUE, styles_xlsx = path, num_styles_csv = num_path, titles = title, footers = footers)
openxlsx::openXL(tab$wb)

# Example 3:  Different pivot
ct <- reshape2::dcast(df, drive  + type ~ colour, value.var= "value", margins=c("drive",  "type"), fun.aggregate = sum)
ct <- ct %>% dplyr::arrange(-row_number())
wb <- xltabr::auto_crosstab_to_wb(ct, titles = title)
openxlsx::openXL(wb)

# Example 4:
ct <- reshape2::dcast(df, drive  + type ~ colour, value.var= "value", margins=c("drive",  "type"), fun.aggregate = sum)
ct <- ct %>% dplyr::arrange(-row_number())
wb <- xltabr::auto_crosstab_to_wb(ct, indent = FALSE, titles = title, footers = footers)
openxlsx::openXL(wb)
