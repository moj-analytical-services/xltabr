# Combined.R contains functions that combine outputs from each
# of the elements (title, top_headers, body etc.)

#' Create all the required properties of the extent
extent_initialise <- function(tab, topleft_row, topleft_col) {

  tab$extent <- list()
  tab$extent$topleft_row <- topleft_row # The top left row we're writing to on the wb
  tab$extent$topleft_col <- topleft_col # The top left col we're writing to on the wb
  tab

}

# Get the rows occupied by the tab
extent_get_rows <- function(tab) {

  tr <- title_get_wb_rows(tab)
  thr <- top_headers_get_wb_rows(tab)
  br <- body_get_wb_rows(tab)
  fr <- footer_get_wb_rows(tab)

  c(tr, thr, br, fr)
}

#Get the cols occupied by the tab
extent_get_cols <- function(tab) {


  tc <- title_get_wb_cols(tab)
  thc <- top_headers_get_wb_cols(tab)
  bc <- body_get_wb_cols(tab)
  fc <- footer_get_wb_cols(tab)

  unique(c(tc, thc, bc, fc))

}


extent_get_bottom_wb_row <- function(tab) {

  er <- extent_get_rows(tab)

  if (length(er)==0) {
    return(tab$extent$topleft_row - 1)
  } else {
    return(max(er))
  }

}

extent_get_rightmost_wb_col <- function(tab) {

  ec <- extent_get_cols(tab)

  if (length(ec)==0) {
    return(tab$extent$topleft_col - 1)
  } else {
    return(max(ec))
  }

}

write_all_elements_to_wb <- function(tab) {
  tab <- xltabr:::title_write_rows(tab)
  tab <- xltabr:::top_headers_write_rows(tab)
  tab <- xltabr:::body_write_rows(tab)
  tab <- xltabr:::footer_write_rows(tab)
  tab
}

#' Get all the styles currently in use by the workbook
combine_all_styles <- function(tab) {

  t1 <- title_get_cell_styles_table(tab)
  t2 <- top_headers_get_cell_styles_table(tab)
  t3 <- body_get_cell_styles_table(tab)
  t4 <- footer_get_cell_styles_table(tab)

  rbind(t1, t2, t3, t4)
}
