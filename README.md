[![Coverage Status](https://img.shields.io/codecov/c/github/moj-analytical-services/xltabr/master.svg)](https://codecov.io/github/moj-analytical-services/xltabr?branch=master) [![Build Status](https://travis-ci.org/moj-analytical-services/xltabr.svg?branch=dev)](https://travis-ci.org/moj-analytical-services/xltabr) [![Cran Status](http://www.r-pkg.org/badges/version/xltabr)](https://cran.r-project.org/web/packages/xltabr/index.html) [![Cran Downloads](https://cranlogs.r-pkg.org/badges/xltabr)](https://www.r-pkg.org/pkg/tidyxl)

**Warning: `xltabr` is in early development. Please raise an [issue](https://github.com/moj-analytical-services/xltabr/issues) if you find any bugs**

Introduction
------------

`xltabr` allows you to write formatted cross tabulations to Excel using [`openxlsx`](https://github.com/awalker89/openxlsx). It has been developed to help automate the process of publishing Official Statistics.

The package works best when the input dataframe is the output of a crosstabulation performed by `reshape2:dcast`. This allows the package to autorecognise various elements of the cross tabulation, which can be styled accordingly.

For example, given a crosstabulation `ct` produced by `reshape2`, the following code produces the table shown.

``` r
titles = c("Breakdown of car statistics", "Cross tabulation of drive and age against type*")
footers = "*age as of January 2015"
wb <- xltabr::auto_crosstab_to_wb(ct, titles = titles, footers = footers)
openxlsx::openXL(wb)
```

![image](vignettes/example_1.png?raw=true)

This readme provides a variety of examples of increasing complexity. It is based on a simulated dataset built into the package, which you can see [here](https://github.com/moj-analytical-services/xltabr/blob/master/inst/extdata/synthetic_data.csv).

Getting started
---------------

Much of `xltabr` utility comes from its ability to automatically format cross tabulations which have been produced by `reshape2:dcast`.

The package provides a core convenience function called `xltabr::auto_crosstab_to_xl`. This wraps more advanced functionality, at the cost of reducing flexibility.

The following code assumes you've read in the synthetic data as follows:

``` r
# Read in data 
path <- system.file("extdata", "synthetic_data.csv", package="xltabr")
df <- read.csv(path, stringsAsFactors = FALSE)
```

### Example 1: Simple cross tabulation to Excel

``` r
# Create a cross tabulation using reshape2
ct <- reshape2::dcast(df, drive + age  ~ type, value.var= "value", margins=c("drive", "age"), fun.aggregate = sum)
ct <- dplyr::arrange(ct, -row_number())

# Use the main convenience function from xltabr to output to excel
tab <- xltabr::auto_crosstab_to_wb(ct, return_tab = TRUE)  #wb is an openxlsx workbook object
openxlsx::openXL(tab$wb)
```

![image](vignettes/example_2.png?raw=true)

### Example 2: Standard data frame to Excel

There is also a convenience function to write a standard data.frame to Excel:

``` r
wb <- xltabr::auto_df_to_wb(mtcars)
openxlsx::openXL(wb)
```

![image](vignettes/example_3.png?raw=true)

### Example 3: Add in titles and footers

``` r
titles = c("Breakdown of car statistics", "Cross tabulation of drive and age against type*")
footers = "*age as of January 2015"
wb <- xltabr::auto_crosstab_to_wb(ct, titles = titles, footers = footers)
openxlsx::openXL(wb)
```

![image](vignettes/example_1.png?raw=true)

### Example 4: Supply custom styles

``` r
path <- system.file("extdata", "styles_pub.xlsx", package = "xltabr")
cell_path <- system.file("extdata", "style_to_excel_number_format_alt.csv", package = "xltabr")
xltabr::set_style_path(path)
xltabr::set_cell_format_path(cell_path)
wb <- xltabr::auto_crosstab_to_wb(ct)
openxlsx::openXL(wb)
```

![image](vignettes/example_5.png?raw=true)

### Example 5: Output more than one table

``` r
# Change back to default styles
xltabr::set_style_path()
xltabr::set_cell_format_path()

# Create second crosstab
ct2 <- reshape2::dcast(df, drive + age ~ colour, value.var= "value", margins=c("drive", "age"), fun.aggregate = sum)
ct2 <- dplyr::arrange(ct2, -row_number())

tab <- xltabr::auto_crosstab_to_wb(ct, titles = titles, footers = c(footers, ""), return_tab = TRUE)

titles2 = c("Table 2: More car statistics", "Cross tabulation of drive and age against colour*")
footers2 = "*age as of January 2015"
wb <- xltabr::auto_crosstab_to_wb(ct2, titles = titles2, footers = footers2, insert_below_tab = tab)
openxlsx::openXL(wb)
```

![image](vignettes/example_4.png?raw=true)

### Example 6: Output more than one table, with different styles

``` r
tab <- xltabr::auto_crosstab_to_wb(ct, titles = titles, footers = c(footers, ""), return_tab = TRUE)

xltabr::set_style_path(path)
xltabr::set_cell_format_path(cell_path)

wb <- xltabr::auto_crosstab_to_wb(ct2, titles = titles2, footers = footers2, insert_below_tab = tab)
openxlsx::openXL(wb)

# Change back to default styles
xltabr::set_style_path()
xltabr::set_cell_format_path()
```

![image](vignettes/example_6.png?raw=true)

### Example 7: Auoindent off

``` r
ct <- reshape2::dcast(df, drive + age  ~ type, value.var= "value",  fun.aggregate = sum)
wb <- xltabr::auto_crosstab_to_wb(ct, titles = titles, footers = c(footers, ""), indent = FALSE, left_header_colnames = c("drive", "age"))
openxlsx::openXL(wb)
```

![image](vignettes/example_7.png?raw=true)

### auto\_crosstab\_to\_wb options

The following provides a list of all the options you can provide to `auto_crosstab_to_wb`

    ## Take a cross tabulation produced by 'reshape2::dcast' and output a formatted openxlsx wb object
    ## 
    ## Description:
    ## 
    ##      Take a cross tabulation produced by 'reshape2::dcast' and output a formatted openxlsx wb object
    ## 
    ## Usage:
    ## 
    ##      auto_crosstab_to_wb(df, auto_number_format = TRUE, top_headers = NULL,
    ##        titles = NULL, footers = NULL, auto_open = FALSE, indent = TRUE,
    ##        left_header_colnames = NULL, vertical_border = TRUE, return_tab = FALSE,
    ##        auto_merge = TRUE, insert_below_tab = NULL, total_text = NULL,
    ##        include_header_rows = TRUE, wb = NULL, ws_name = NULL,
    ##        number_format_overrides = list(), fill_non_values_with = list(na = NULL,
    ##        nan = NULL, inf = NULL, neg_inf = NULL), allcount_to_level_translate = NULL,
    ##        left_header_col_widths = NULL, body_header_col_widths = NULL)
    ##      
    ## Arguments:
    ## 
    ##       df: A data.frame.  The cross tabulation to convert to Excel
    ## 
    ## auto_number_format: Whether to automatically detect number format
    ## 
    ## top_headers: A list.  Custom top headers. See 'add_top_headers()'
    ## 
    ##   titles: The title.  A character vector.  One element per row of title
    ## 
    ##  footers: Table footers.  A character vector.  One element per row of footer.
    ## 
    ## auto_open: Boolean. Automatically open Excel output.
    ## 
    ##   indent: Automatically detect level of indentation of each row
    ## 
    ## left_header_colnames: The names of the columns that you want to designate as left headers
    ## 
    ## vertical_border: Boolean. Do you want a left border?
    ## 
    ## return_tab: Boolean.  Return a tab object rather than a openxlsx workbook object
    ## 
    ## auto_merge: Boolean.  Whether to merge cells in the title and footers to width of body
    ## 
    ## insert_below_tab: A existing tab object.  If provided, this table will be written on the same sheet, below the provided tab.
    ## 
    ## total_text: The text that is used for the 'grand total' of a cross tabulation
    ## 
    ## include_header_rows: Boolean - whether to include or omit the header rows
    ## 
    ##       wb: A existing openxlsx workbook.  If not provided, a new one will be created
    ## 
    ##  ws_name: The name of the worksheet you want to write to
    ## 
    ## number_format_overrides: e.g. list("colname1" = "currency1") see auto_style_number_formatting
    ## 
    ## fill_non_values_with: Manually specify a list of strings that will replace non numbers types NA, NaN, Inf and -Inf. e.g. list(na = '*', nan = '', inf = '-', neg_inf = '-'). Note: NaNs are not treated as NAs.
    ## 
    ## allcount_to_level_translate: Manually specify how to translate summary levels into header formatting
    ## 
    ## left_header_col_widths: Width of row header columns you wish to set in Excel column width units. If singular, value is applied to all row header columns. If a vector, vector must have length equal to the number of row headers in workbook. Use special case "auto" for
    ##           automatic sizing. Default (NULL) leaves column widths unchanged.
    ## 
    ## body_header_col_widths: Width of body header columns you wish to set in Excel column width units. If singular, value is applied to all body columns. If a vector, vector must have length equal to the number of body headers in workbook. Use special case "auto" for
    ##           automatic sizing. Default (NULL) leaves column widths unchanged.
    ## 
    ## Examples:
    ## 
    ##      crosstab <- read.csv(system.file("extdata", "example_crosstab.csv", package="xltabr"))
    ##      wb <- auto_crosstab_to_wb(crosstab)
    ## 

Advanced usage
--------------

The simple examples above wrap lower-level functions. These functions can be used to customise the output in a number of ways.

The following example shows the range of functions available.

``` r
tab <- xltabr::initialise() %>%  #Options here for providing an existing workbook, changing worksheet name, and position of table in wb
  xltabr::add_title(title_text) %>% # Optional title_style_names allows user to specify formatting
  xltabr::add_top_headers(h_list) %>% # Optional row_style_names and col_style_names allows custom formatting
  xltabr::add_body(df) %>%  #Optional left_header_colnames, row_style_names, left_header_style_names col_style names
  xltabr::add_footer(footer_text) %>% # Optional footer_style_names
  xltabr::auto_detect_left_headers() %>% # Auto detect left headers through presence of keyword, default = '(all)'
  xltabr::auto_detect_body_title_level() %>% # Auto detect level of emphasis of each row in body, through presence of keyword
  xltabr::auto_style_indent() %>% # Consolidate all left headers into a single column, with indentation to signify emphasis level
  xltabr::auto_merge_title_cells() %>% # merge the cells in the title
  xltabr::auto_merge_footer_cells() # merge the cells in the footer
```

The convenience functions contain further examples of how to build up a tab. See [here](https://github.com/moj-analytical-services/xltabr/blob/dev/R/convenience.R).

Implementation diagrams.
------------------------

See [here](https://www.draw.io/?lightbox=1&highlight=0000ff&edit=_blank&layers=1&nav=1&title=xltabr#Uhttps%3A%2F%2Fdrive.google.com%2Fa%2Fdigital.justice.gov.uk%2Fuc%3Fid%3D0BwYwuy7YhhdxY2hGQnVGNFN6QkE%26export%3Ddownload)
