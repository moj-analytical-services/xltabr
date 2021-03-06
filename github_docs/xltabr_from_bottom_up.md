Introduction
------------

This documentation explains the internals of xltabr, and is intended for advanced users who want greater control of the formatting of their outputs.

You may find [this diagram](https://www.draw.io/?lightbox=1&highlight=0000ff&edit=_blank&layers=1&nav=1&title=xltabr#Uhttps%3A%2F%2Fdrive.google.com%2Fa%2Fdigital.justice.gov.uk%2Fuc%3Fid%3D0BwYwuy7YhhdxY2hGQnVGNFN6QkE%26export%3Ddownload) helpful as you work through this document.

The main `tab` object
---------------------

Users of the xltabr package progressively build up an R list which, by convention, is called `tab`. For instance, we add titles to this object using `xltabr::add_titles(tab)` which adds information to the list at `tab$titles`. Similarly `xltabr::add_body(tab)` adds information to `tab$body`. Other general information is stored in additional keys like `xltabr$style_catalogue` and `xltabr$misc`.

The `tab` object contains an `openxlsx` workbook object at `tab$wb`. Once the user has added all the elements and styling to the tab object, the user calls `xltabr::write_data_and_styles_to_wb`, which updates the workbook at `tab$wb`. The user can then open this in Excel or save it using `openxlsx::openXL(tab$wb)` or `openxlsx::saveWorkbook(tab$wb)`.

A good place to start in understanding `xltabr` is to generate a `tab` object and start exploring its contents. which you can do using the following code:

``` r
path <- system.file("extdata", "synthetic_data.csv", package="xltabr")
df <- read.csv(path, stringsAsFactors = FALSE)
ct <- reshape2::dcast(df, drive + age  ~ type, value.var= "value", margins=c("drive", "age"), fun.aggregate = sum)
ct <- dplyr::arrange(ct, -row_number())
tab <- xltabr::auto_crosstab_to_wb(ct, return_tab = TRUE) # this is a convenience function that calls a number of xltabr functions
View(tab)
```

Building blocks
---------------

Styles are read into xltabr from an Excel style sheet. The user defines the styles in an Excel spreadsheet, and reads these styles into xltabr. Styles can then be applied to cells in Excel output tables. The default stylesheet can be found [here](https://github.com/moj-analytical-services/xltabr/blob/master/inst/extdata/styles.xlsx) in the repo, and you can download it [here](https://github.com/moj-analytical-services/xltabr/blob/master/inst/extdata/styles.xlsx?raw=true).

When styles are read in, they are created as `xltabr_style`s, the definition for which can be found [here](https://github.com/moj-analytical-services/xltabr/blob/ec84ae6260a0eb4c513e8f097eb316e16e9a6c7c/R/style_catalogue.R#L1).

The Excel stylesheet provides a set of 'base' styles. Styles will later be combined, in a way that borrows concepts from [Cascading Style Sheets (CSS)](https://en.wikipedia.org/wiki/Cascading_Style_Sheets). Specifically, it is common for several styles to be applied to a cell. For instance, we may wish to assign `body|title_1|currency` to a cell. This means apply `body` first, then apply `title_1` on top of body, overriding any styles in `body` which are also defined in `title_1`. Then apply `currency` on top of `body|title_1` in the same way.

You can explore the style catalogue by running the following code:

``` r
# You can use xltabr::set_style_path() to use your own style sheet.  Otherwise the default one is used.
tab <- xltabr::initialise()  # Under the hood this calls xltabr:::style_catalogue_initialise(tab)
View(tab$style_catalogue)
```

Now we have a style catalogue, we need a way of applying these styles to tables of data.

One way of doing this would be to define a style for each individual cell of the output table. However, this would require writing code for each specific cell, which would be too labour-intensive. We can do better by recognising that it's typical for all cells in a column, and all cells in a row, to share similar styling.

Building on this idea, in `xltabr` we define a style for each row and each column of the output table. The 'final style' which is applied is then a combination of the row and column style, where the base style is the row style, which then inherits additional properties from the column style.

``` r
tab <- xltabr::initialise()  # Under the hood this calls xltabr:::style_catalogue_initialise(tab)
View(tab$style_catalogue)
```

Styling a table by hand
-----------------------

Most users will use the various 'automatic styling' functions to avoid having to manually define styles. However, `xltabr` allows the user to do this manually. This can be useful for heavily customised output, or just if the user needs to override the styles for a specific row or column after using the automatic styling functions like `xltabr::auto_crosstab_to_tab`.

Let's consider creating and styling a simple 3x2 table in xltabr. Note the style names are defined in the [default style sheet](https://github.com/moj-analytical-services/xltabr/blob/master/inst/extdata/styles.xlsx). The style names of number formats are defined [here](https://github.com/moj-analytical-services/xltabr/blob/master/inst/extdata/number_format_defaults.csv).

``` r
df <-  data.frame("cat" = c("Apples", "Bananas", "Total"), 
                  "q" = c(15,5,20))

# Initialise the 'tab' object using default params
# This means a new workbook will be created, writing to Sheet1, using the default style sheet
tab <- xltabr::initialise()

# Add top headers and associated styles
top_header_text <- c("Category", "Quantity")
top_header_style <- "top_header_1"
tab <- add_top_headers(tab, top_headers = top_header_text, row_style_names = top_header_style)

# Add the data to the tab object  
row_style_names <- c("body", "body", "body|title_3")
column_style_names <- c("right_border|character","right_border_integer")
tab <- add_body(tab, df, row_style_names = row_style_names, col_style_names = column_style_names)

# Note up to now, we've just been preparing 'meta data' so the tab object knows what it should write to the Excel workbook
# No data has actually been written to the workbook itself

# Write out styles and data to workbook
tab <- write_data_and_styles_to_wb(tab)

openxlsx::openXL(tab$wb)
```

We can gain a deeper understanding of the `xltabr` by investigating how this information is stored in the `tab` object. The styling information for the top headers (the column titles) is stored in \`tab$top\_headers

``` r
tab$top_headers$top_headers_list
```

    ## [[1]]
    ## [1] "Category" "Quantity"

``` r
tab$top_headers$top_headers_row_style_names
```

    ## row_style_names 
    ##  "top_header_1"

We can see that table data is stored in the `body_df_to_write`:

``` r
tab$body$body_df_to_write
```

    ##       cat  q
    ## 1  Apples 15
    ## 2 Bananas  5
    ## 3   Total 20

And that additional meta data is stored in `body_df` (ignore 'meta\_left\_header\_row\_' for now)

``` r
tab$body$body_df
```

    ##       cat  q    meta_row_ meta_left_header_row_
    ## 1  Apples 15         body      body|left_header
    ## 2 Bananas  5         body      body|left_header
    ## 3   Total 20 body|title_3      body|left_header

Finally, column styling information for the body is stored here:

``` r
tab$body$meta_col_
```

    ## [1] "right_border|character" "right_border_integer"

Left headers and
----------------
