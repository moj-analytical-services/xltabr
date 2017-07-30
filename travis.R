packages <- c("rmarkdown")
new_packages <-  packages[!(packages %in% installed.packages()[,"Package"])]

if (length(new_packages)) {
  install.packages(new_packages, repos = "http://cran.us.r-project.org")
}


library(rmarkdown)
rmarkdown::render("vignettes/readme.Rmd", output_format = "md_document", output_file = "../README.md", output_options = list("variant" = "markdown_github"))
