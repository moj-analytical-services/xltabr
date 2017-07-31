[![Coverage Status](https://img.shields.io/codecov/c/github/moj-analytical-services/xltabr/dev.svg)](https://codecov.io/github/moj-analytical-services/xltabr?branch=dev) [![Build Status](https://travis-ci.org/moj-analytical-services/xltabr.svg?branch=dev)](https://travis-ci.org/moj-analytical-services/xltabr)

Introduction
------------

xltabr allows you to write formatted cross tabulations to Excel using [`openxlsx`](https://github.com/awalker89/openxlsx). It has been developed to help automate the process of publishing official statistics.

The user provides a dataframe, which is outputted to Excel with various types of rich formatting such as totals and subtotals shown with various degrees of emphasis, number formatting, etc.

The package works best when the input dataframe is the output of a crosstabulation performed by `reshape2:dcast`. This allows the package to autorecognise key elements of a cross tabulation, such as which element form the body, and which elements form the headers.

Getting started
---------------

Much of `xltabr` utility comes from its ability to automatically format cross tabulations which have been produced by `reshape2:dcast`.

The package provides a core convenience function called `xltabr::auto_crosstab_to_xl`. This wraps more advanced functionality, at the cost of reducing flexibility.

A simple example is as follows

``` r
ct <- reshape2::dcast(mtcars, am + gear ~ cyl, value.var= "mpg", margins=c("am", "gear"), fun.aggregate = mean)
wb <- xltabr::auto_crosstab_to_xl(ct)
openxlsx::openXL(wb)
```

The simplest use is to write an aribtary data.frame to an in memory openxlsx workbook, ready to be written to disk:

``` r
wb <- xltabr::auto_df_to_xl(mtcars)
openxlsx::openXL(wb)
```

### Options

    ## Take a cross tabulation produced by 'reshape2::dcast' and output a
    ## formatted openxlsx wb objet
    ## 
    ## Description:
    ## 
    ##      Take a cross tabulation produced by `reshape2::dcast` and output a
    ##      formatted openxlsx wb objet
    ## 
    ## Usage:
    ## 
    ##      auto_crosstab_to_wb(df, auto_number_format = TRUE, top_headers = NULL,
    ##        titles = NULL, footers = NULL, auto_open = FALSE, indent = TRUE,
    ##        left_header_colnames = NULL, return_tab = FALSE)
    ##      
    ## Arguments:
    ## 
    ##       df: A data.frame.  The cross tabulation to convert to Excel
    ## 
    ## auto_number_format: Whether to automatically detect number format
    ## 
    ## auto_open: Automatically open Excel?
    ## 
    ## top_headers.: A list.  Custom top headers.
    ## 
    ##    title: The title.  A character vector.  One element per row of title
    ## 
    ##   footer: Table footers.  A character vector.  One element per row of
    ##           footer.

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
  xltabr:::auto_detect_left_headers() %>% # Auto detect left headers through presence of keyword, default = '(all)'
  xltabr:::auto_detect_body_title_level() %>% # Auto detect level of emphasis of each row in body, through presence of keyword
  xltabr:::auto_style_indent() # Consolidate all left headers into a single column, with indentation to signify emphasis level
```

### Additional style customisation

The default styles can be found in an Excel workbook bundled with the package, at `system.file("extdata", "styles.xlsx", package="xltabr")` or [here](https://github.com/moj-analytical-services/xltabr/blob/dev/inst/extdata/styles.xlsx).

Some diagrams
-------------

See [here](https://drive.google.com/file/d/0BwYwuy7YhhdxY2hGQnVGNFN6QkE/view?usp=sharing)

This documentation was automatically generated on 31 July, 2017 by [Travis CI](https://travis-ci.org/).
