sheetExists <- function(wb, ws_name) {
  sheets <- openxlsx::sheets(wb)
  return(ws_name %in% sheets)
}

df_width <- function(tab) {
  ncol(tab$data$df_final)
}


get_style <- function(tab, row, col) {

  filtered <- tab$styles_lookup %>%
    dplyr::filter(row == row) %>%
    dplyr::filter(col == col) %>%
    select(style)

  if (length(filtered) == 1 ) {
    return(filtered[1,1])
  } else {
    return(NULL)
  }



}

# Takes a style object and updates it based on a se
update_style <- function(style_old, style_new) {

}

# Unfortunately it seems that finding the currently styling of a particular row and column is quite difficult
# Create a df with a list column that contains all the current styles for all the worksheets that can then easily be lookuped up
update_styles_lookup <- function(tab) {

  stobs <- tab$wb$styleObjects

  counter <- 1

  style <- list()
  col <- integer()
  row <- integer()

  for (stob in stobs) {
    for (i in seq_along(stob$rows)) {
      row[counter] <- stob$rows[i]
      col[counter] <- stob$cols[i]
      style[counter] <- stob$style
      counter = counter + 1
    }
  }

  tab$styles_lookup <- dplyr::data_frame(col = col, row = row, style = style)

}
