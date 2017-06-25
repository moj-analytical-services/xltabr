#' Create all the required properties of the extent
extent_initialise <- function(tab, topleft_row, topleft_col) {

  tab$extent <- list()
  tab$extent$topleft_row <- topleft_row # The top left row we're writing to on the wb
  tab$extent$topleft_col <- topleft_col # The top left col we're writing to on the wb
  tab

}

# Get the rows occupied by the tab
extent_get_rows <- function(tab) {

  tr <- title_get_rows(tab)
  tr
}

#Get the cols occupied by the tab
extent_get_cols <- function(tab) {

  tc <- title_get_cols(tab)
  tc

}


extent_get_bottomright_row <- function(tab) {

  er <- extent_get_rows(tab)

  if (length(er)==0) {
    return(tab$extent$topleft_row - 1)
  } else {
    return(max(er))
  }

}

extent_get_bottomright_col <- function(tab) {

  ec <- extent_get_cols(tab)

  if (length(ec)==0) {
    return(tab$extent$topleft_col - 1)
  } else {
    return(max(ec))
  }

}
