library(magrittr)

title_text <- c("Table 1.1: Prison population by type of custody, age group and sex", "Subtitle text goes here", "")
title_style_names <- c("title", "subtitle", "spacer")

tab <- xltabr::initialise() %>%
  xltabr::add_title(title_text, title_style_names)

xltabr:::title_get_style_names(tab)
