# Use default .xlsx style catalogue to initialise the catalogue of styles
# Following initialisation, we have tab$style catalogue, which contains each style.
# Values are chracter 'structures' e.g. '"structure(list(valign = \"center\"), .Names = \"valign\")"'
style_catalogue_initialise <- function(tab) {

  tab <- style_catalogue_xlsx_import(tab)
  tab <- style_catalogue_import_num_formats(tab)
  tab
}

# Number formats are stored in a csv, by default at xltabr/inst/extdata/style_to_excel_number_format.csv
# This is a named style (e.g. percent1) whcich translates into an Excel style (e.g. #,0.0%;#,0.0%;;* @)
style_catalogue_import_num_formats <- function(tab){

  lookup_df <- utils::read.csv(get_cell_format_path(), stringsAsFactors = FALSE)

  # Convert dataframe into two vectors
  style_keys <- lookup_df$excel_format
  style_names <- lookup_df$style_name

  if (length(style_names) == 0) {
    warning("Your number format excel contained no data.  For an example see xltabr/inst/extdata/style_to_excel_number_format.csv")
    return(tab)
  }

  for (i in 1:length(style_names)) {
    tab$style_catalogue[[style_names[i]]] <- convert_style_list_to_character_key(list(numFmt = style_keys[i]))
  }
  tab
}

# Iterate through the default style workbook, adding them to the styles catalogue
style_catalogue_xlsx_import <- function(tab) {

  path <- get_style_path()

  wb <- openxlsx::loadWorkbook(path)

  # When you load an Excel document, openxlsx creates a list of the distinct style objects in the workbook
  # Iterate through these style objects
  for (this_style_object in wb$styleObjects) {

    # It's possible that two styles (e.g. top header, title1) might have the same style.  If so, create an entry for both
    for (this_row in 1:length(this_style_object$rows)) {

      r <- this_style_object$rows[this_row]
      c <- 1 # We assume that all styles are stored in the leftmost column

      suppressWarnings(cell <- openxlsx::readWorkbook(wb, rows = r, cols = c, colNames = FALSE, rowNames = FALSE))

      style_cat_name <- cell[1, 1]

      if (is_null_or_blank(style_cat_name)) {
        next
      }

      #See https://github.com/moj-analytical-services/xltabr/issues/88
      style_cat_name <- as.character(style_cat_name)

      # Convert from s4 object to list
      style_list <- convert_style_object_to_list(this_style_object$style)

      # Add row height property

      #Convert from list to charaster string
      style_key <- convert_style_list_to_character_key(style_list)

      tab$style_catalogue[[style_cat_name]] <- style_key
    }
  }

  tab
}

# Serialises a list to a character string
# Returns a style_key string based on our style catalogue objects (should also work with default R lists)
convert_style_list_to_character_key <- function(style_list){
  style_key <- gsub(' +', ' ', paste0(utils::capture.output(dput(style_list)), collapse = ""))
  style_key
}


# Parses a character style_key to style_list
style_key_parser <- function(style_key){
  style_list <- eval(parse(text = style_key))
  style_list
}

# Where we have an existing style, and a new style, this function determines the inheritance rules
# similar to cascading style sheets
# e.g. in body|header_1, we want header_1 styles to over
cascade_style <- function(existing_style, new_style) {

  for (property_name in names(new_style)) {

    # fontDectoration is a special case because we we want bold + italic to = bolditalic, not italic
    if (property_name == "fontDecoration") {
      existing_style[[property_name]] <- unique(c(existing_style[[property_name]], new_style[[property_name]]))
    } else {
      existing_style[property_name] <- new_style[property_name]
    }
  }

  existing_style

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
  if (!all(style_check)) {
    stop(paste("The following style names:",  paste0(ussd[!style_check], collapse = ", "), "are not in the style_catalogue please add then maunally or specify them in style.xlsx"))
  }


  if (length(seperated_style_definition) <= 1) {
    return(tab$style_catalogue[[seperated_style_definition]])
  }
  else{
    # Otherwise build final style
    previous_style <- style_key_parser(tab$style_catalogue[[seperated_style_definition[1]]])
    for (base_style_index in 2:length(seperated_style_definition)) {
      current_style <- style_key_parser(tab$style_catalogue[[seperated_style_definition[base_style_index]]])
      previous_style <- cascade_style(previous_style, current_style)
    }
  }
  return(convert_style_list_to_character_key(previous_style))
}

add_style_defintions_to_catalogue <- function(tab, style_definitions){

  for (style_def in style_definitions) {
    style_key <- build_style(tab, style_def)

    # Always log the key value pair in the style dictionary
    if (!style_def %in% names(tab$style_catalogue)) {
      tab$style_catalogue[[style_def]] <- style_key
    }
  }

  tab
}

convert_style_object_to_list <- function(style, convert_to_S4 = FALSE){
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

  # Get a table with a style for each cell
  full_table <- combine_all_styles(tab)

  if (is.null(full_table)) stop("Please ensure add_from has appropriate values and a table has been added to tab")

  # Bloody factors
  full_table <- data.frame(lapply(full_table, as.character), stringsAsFactors = FALSE)

  # Get a unique vector of style_name from full_table
  unique_styles_definitions <- unique(full_table$style_name)

  # Add unique style_name vector to style catalogue
  tab <- add_style_defintions_to_catalogue(tab, unique_styles_definitions)

  # add the style_key to each style name in full table
  full_table$style_key <- unlist(tab$style_catalogue[full_table$style_name])

  # iterate over a unique list of style_keys for and apply them to the workbook for each row col
  unique_style_keys <- unique(full_table$style_key)

  for (sk in unique_style_keys){
    row_col_styles <- full_table[full_table$style_key == sk,]
    rows <- row_col_styles$row
    cols <- row_col_styles$col

    created_style <- convert_style_object_to_list(sk, convert_to_S4 = TRUE)

    # Add cell style
    openxlsx::addStyle(tab$wb, tab$misc$ws_name, created_style, rows, cols)

    # If this row's row height is greater than Apply row height.

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
