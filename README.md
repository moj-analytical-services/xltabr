[![Coverage Status](https://img.shields.io/codecov/c/github/moj-analytical-services/xltabr/dev.svg)](https://codecov.io/github/moj-analytical-services/xltabr?branch=dev) [![Build Status](https://travis-ci.org/moj-analytical-services/xltabr.svg?branch=dev)](https://travis-ci.org/moj-analytical-services/xltabr)

Introduction
------------

xltabr allows you to write formatted cross tabulations to Excel usign openxlsx. It has been developed to help automate the process of publishing official statistics.

The user provides a dataframe, which is outputted to Excel with various types of rich formatting such as totals and subtotals shown with various degrees of emphasis, number formatting, etc.

The package works best when the input dataframe is the output of a crosstabulation performed by `reshape2:dcast`. This allows the package to autorecognise key elements of a cross tabulation, such as which element form the body, and which elements form the headers.

Getting started
---------------

TODO: Auto generate this readme from a rmd

Some diagrams
-------------

See [here](https://drive.google.com/file/d/0BwYwuy7YhhdxY2hGQnVGNFN6QkE/view?usp=sharing)

rmarkdown::render("vignettes/readme.Rmd", output\_format="md\_document", output\_file = "../README.md", output\_options=list("variant" = "markdown\_github"))
