# Use default .xlsx style catalogue to initialise the catalogue of styles
style_catalogue_initialise <- function(tab, styles_xlsx = NULL, num_styles_csv = NULL) {
  if (is.null(styles_xlsx)) {
    path <- system.file("extdata", "styles.xlsx", package = "xltabr")
  } else {
    path <- styles_xlsx
  }

  if (is.null(num_styles_csv)) {
    path_num <- system.file("extdata", "style_to_excel_number_format.csv", package = "xltabr")
  } else {
    path_num <- num_styles_csv
  }
  tab <- style_catalogue_xlsx_import(tab, path)
  tab <- style_catalogue_import_num_formats(tab, path_num)
  tab
}

style_catalogue_import_num_formats <- function(tab, path){
  # This lookup table coverts

  lookup_df <- utils::read.csv(path, stringsAsFactors = FALSE)

  # Convert dataframe into two vectors
  style_keys <- lookup_df$excel_format
  style_names <- lookup_df$style_name

  for (i in 1:length(style_names)){
    tab$style_catalogue[[style_names[i]]] <- create_style_key(list(numFmt = style_keys[i]))
  }
  tab
}

# Iterate through the default style workbook, adding them to the styles catalogue
style_catalogue_xlsx_import <- function(tab, path) {

  # if initialising style_catalogue do we want to reset it with line below?
  tab$style_catalogue <- list()

  wb <- openxlsx::loadWorkbook(path)
  listed_styles <- openxlsx::readWorkbook(wb, colNames = FALSE)

  for (i in wb$styleObjects) {
    for (iter in 1:length(i$rows)){
      r <- i$rows[iter]
      c <- i$cols[iter]
      suppressWarnings(cell <- openxlsx::readWorkbook(wb, rows = r, cols = c, colNames = FALSE, rowNames = FALSE))

      value <- cell[1, 1]

      if (is.null(value)) {
        next
      }

      # tmp_list <- list()
      # tmp_list$style <- i$style

      # cell <- openxlsx::readWorkbook(wb, rows = r, cols = c + 1, colNames = FALSE, rowNames = FALSE)
      # tmp_list$rowHeight <- cell[1, 1]

      style_list <- convert_style_object(i$style)
      style_key <- create_style_key(style_list)

      if (!value %in% names(tab$style_catalogue)){
        tab$style_catalogue[[value]] <- style_key
      }
    }
  }

  # Add in a checker to catch default style objects
  for (style_name in listed_styles$X1){
    if (!style_name %in% names(tab$style_catalogue)){
      tab$style_catalogue[[style_name]] <- create_style_key(convert_style_object(openxlsx::createStyle()))
    }
  }

  tab
}

#' Allows the user to provide a .xlsx with custom defined styles
#' Use [this](https://github.com/moj-analytical-services/xltabr/blob/dev/inst/extdata/styles.xlsx?raw=true) file as a basis
#'
#' @param path a file path to the .xlsx file you want to use to override styles
#'
#' @return The tab
#' @export
style_catalogue_override_styles <- function(tab, path_to_xlsx) {
  # Code here to override styles from Excel document provided by user
  tab <- style_catalogue_xlsx_import(tab, path_to_xlsx)
  tab
}

# # # # # Added functions by Karik # # # # #

# # # # # # # # # # # # # # # # # # # #

# add_to_dictionary
# returns a style_key string based on our style catelogue objects (should also work with default R lists)
create_style_key <- function(style_list){
  style_key <- gsub(' +', ' ', paste0(utils::capture.output(dput(style_list)), collapse = ""))

  style_key
}

# Converts style object property to string representation (used in create_style_key function) Depreciated
property_to_key <- function(style_object, property){
  if(is.null(style_object[[property]])){
    return(NULL)
  }
  else{
    return(paste(property, paste0(style_object[[property]], collapse = "%"), sep = "_"))
  }
}

# Converts style_key to style_list
style_key_parser <- function(style_key){
  style_list <- eval(parse(text=style_key))
  style_list
}

# looks at the cell_style_definition and build a final style_key for that cell
build_style <- function(tab, cell_style_definition){

  # Convert cell style inheritence string into an array
  seperated_style_definition <- unlist(strsplit(cell_style_definition, "\\|"))

  ## Run a check (that all base styles referenced in cell_style_definition exist in style_catalogue)
  # Get array of base styles (i.e. styles that are not a combination of multiple styles (no pipes in names))
  base_styles <- names(tab$style_catalogue)[!grepl("\\|",names(tab$style_catalogue))]
  ussd <- unique(seperated_style_definition)
  style_check <- ussd %in% base_styles
  if(!all(style_check)){
    stop(paste("The following style names:",  paste0(ussd[!style_check], collapse = ", "), "are not in the style_catalogue please add then maunally or specify them in style.xlsx"))
  }
  ##

  if (length(seperated_style_definition) <= 1){
    return (tab$style_catalogue[[seperated_style_definition]])
  }
  else{
    # Otherwise build final style
    previous_style <- style_key_parser(tab$style_catalogue[[seperated_style_definition[1]]])
    for (i in 2:length(seperated_style_definition)){
      current_style <- style_key_parser(tab$style_catalogue[[seperated_style_definition[i]]])
      for (property_name in names(current_style)){
        if(property_name == "fontDecoration"){
          previous_style[[property_name]] <- unique(c(previous_style[[property_name]], current_style[[property_name]]))
        } else {
          previous_style[property_name] <- current_style[property_name]
        }
      }
    }
  }
  return (create_style_key(previous_style))
}

add_style_defintions_to_catelogue <- function(tab, style_definitions){

  for (style_def in style_definitions){
    style_key <- build_style(tab, style_def)

    # Always log the key value pair in the style dictionary
    if (!style_def %in% names(tab$style_catalogue)){
      tab$style_catalogue[[style_def]] <- style_key
    }
  }

  tab
}

convert_style_object <- function(style, convert_to_S4 = FALSE){
  if(convert_to_S4){

    if(typeof(style) == "character"){
      style <- style_key_parser(style)
    }

    style_properties <- names(style)

    if("numFmt" %in% style_properties){
      style_properties <- style_properties[!("numFmt" == style_properties)]
      out_style <- openxlsx::createStyle(numFmt = style[["numFmt"]])
    } else{
      out_style <- openxlsx::createStyle()
    }

    for (prop in style_properties){
      out_style[[prop]] <- style[[prop]]
    }

    return(out_style)

  } else {
    out_style <- style$as.list()

    # For some reason fills do not export properly so workaround is added here
    if(any(c("fillFg", "fillBg") %in% names(out_style))){
      out_style[["fill"]] <- list(fillFg = out_style[["fillFg"]], fillBg = out_style[["fillBg"]])
      out_style[["fillFg"]] <- NULL
      out_style[["fillBg"]] <- NULL
    }
    return(out_style)
  }
}


add_styles_to_wb <- function(tab){

  full_table <- combine_all_styles(tab)

  if(is.null(full_table)) stop("Please ensure add_from has appropriate values and a table has been added to tab")

  # Bloody factors
  full_table <- data.frame(lapply(full_table, as.character), stringsAsFactors=FALSE)

  # Get a unique vector of style_name from full_table
  unique_styles_definitions <- unique(full_table$style_name)

  # Add unique style_name vector to style catalogue
  tab <- add_style_defintions_to_catelogue(tab, unique_styles_definitions)

  # add the style_key to each style name in full table
  full_table$style_key <- unlist(tab$style_catalogue[full_table$style_name])

  # iterate over a unique list of style_keys for and apply them to the workbook for each row col
  unique_style_keys <- unique(full_table$style_key)

  for (sk in unique_style_keys){
    row_col_styles <- full_table[full_table$style_key == sk,]
    rows <- row_col_styles$row
    cols <- row_col_styles$col

    created_style <- convert_style_object(sk, convert_to_S4 = TRUE)
    openxlsx::addStyle(tab$wb, tab$misc$ws_name, created_style, rows, cols)
  }

  tab
}

# Not used in package - used for debug
compare_style_lists <- function(a, b){
  a_names <- sort(names(a))
  b_names <- sort(names(b))

  if(length(a_names) != length(b_names)) return(FALSE)
  if(!all(a_names == b_names)) return(FALSE)

  final_check <- TRUE
  for (prop in a_names){
    if(!identical(a[prop], b[prop])){
      final_check <- FALSE
      break;
    }
  }
  return(final_check)
}
