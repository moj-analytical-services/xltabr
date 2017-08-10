# Use default .xlsx style catalogue to initialise the catalogue of styles
style_catalogue_initialise <- function (tab) {
  path <- system.file("extdata", "styles.xlsx", package = "xltabr")
  path_num <- system.file("extdata", "style_to_excel_number_format.csv", package = "xltabr")
  tab <- style_catalogue_xlsx_import(tab, path)
  tab <- style_catalogue_import_num_formats(tab, path_num)
  tab
}

style_catalogue_import_num_formats <- function(tab, path){
  # This lookup table coverts

  lookup_df <- read.csv(path, stringsAsFactors = FALSE)

  # Convert dataframe into two vectors
  style_keys <- paste0("numFmt_", lookup_df$excel_format)
  style_names <- lookup_df$style_name

  for (i in 1:length(style_names)){
    tab$style_catalogue[[style_names[i]]] <- style_keys[i]
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
      cell <- openxlsx::readWorkbook(wb, rows = r, cols = c, colNames = FALSE, rowNames = FALSE)

      value <- cell[1, 1]

      tmp_list <- list()
      tmp_list$style <- i$style

      cell <- openxlsx::readWorkbook(wb, rows = r, cols = c + 1, colNames = FALSE, rowNames = FALSE)
      tmp_list$rowHeight <- cell[1, 1]

      style_key <- create_style_key(convert_style_object(tmp_list))

      if (!value %in% names(tab$style_catalogue)){
        tab$style_catalogue[[value]] <- style_key
      }
    }
  }

  # Add in a checker to catch default style objects
  for (style_name in listed_styles$X1){
    if (!style_name %in% names(tab$style_catalogue)){
      style_key <- "numFmt_GENERAL"
      tab$style_catalogue[[style_name]] <- style_key
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

  # NULL returns from property_to_key are automatically removed from vector
  style_properties <- c(property_to_key(style_list, "fontName"),
                        property_to_key(style_list, "fontSize"),
                        property_to_key(style_list, "fontColour"),
                        property_to_key(style_list, "numFmt"),
                        property_to_key(style_list, "border"),
                        property_to_key(style_list, "borderColour"),
                        property_to_key(style_list, "borders"),
                        property_to_key(style_list, "bgFill"),
                        property_to_key(style_list, "fgFill"),
                        property_to_key(style_list, "halign"),
                        property_to_key(style_list, "valign"),
                        property_to_key(style_list, "textDecoration"),
                        property_to_key(style_list, "wrapText"),
                        property_to_key(style_list, "textRotation"),
                        property_to_key(style_list, "indent"),
                        property_to_key(style_list, "rowHeight"))

  style_key <- paste(style_properties, collapse = "|")

  style_key
}

# Converts style object property to string representation (used in create_style_key function)
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

  specs <- unlist(strsplit(style_key, "\\|"))

  # Can interate through style categories in order to get inherited changes from left to right.
  style_categories <- list()
  for (i in 1:length(specs)){
    seperate_property <- unlist(strsplit(specs[[i]], "_"))
    style_categories[seperate_property[1]] <- unlist(strsplit(seperate_property[2], "%"))
  }

  style_categories
}

# looks at the cell_style_definition and build a final style_key for that cell
build_style <- function(cell_style_definition, style_catalogue){

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
    return (style_catalogue[[seperated_style_definition]])
  }
  else{
    # Otherwise build final style
    previous_style <- style_key_parser(style_catalogue[[seperated_style_definition[1]]])
    for (i in 2:length(seperated_style_definition)){
      current_style <- style_key_parser(style_catalogue[[seperated_style_definition[i]]])
      for (property_name in names(current_style)){
        previous_style[property_name] <- current_style[property_name]
      }
    }
  }
  return (create_style_key(previous_style))
}

add_style_defintions_to_catelogue <- function(style_catalogue, style_definitions){

  for (style_def in style_definitions){
    style_key <- build_style(style_def, style_catalogue)

    # Always log the key value pair in the style dictionary
    if (!style_def %in% names(style_catalogue)){
      style_catalogue[[style_def]] <- style_key
    }
  }

  style_catalogue
}

# style_key <- tab$style_catalogue[["body|indent_0"]]
# style <- style_key
# convert_to_S4 = T

convert_style_object <- function(style, convert_to_S4 = FALSE){
  if(convert_to_S4){

    if(typeof(style) == "character"){
      style <- style_key_parser(style)
    }

    # Do convertions for correct createStyle inputs
    font_colour_input <- NULL
    if(!is.null(style[["fontColour"]])){
      # if(style[["fontColour"]] "1") style[["fontColour"]] <- "black"
      if(suppressWarnings(is.na(as.integer(style[["fontColour"]])))){
        font_colour_input <- style[["fontColour"]]
      }
    }

    if(!is.null(style[["fontSize"]])){
      style[["fontSize"]] <- as.integer(style[["fontSize"]])
    }

    if(!is.null(style[["indent"]])){
      style[["indent"]] <- as.integer(style[["indent"]])
    }

    if(is.null(style[["numFmt"]])){
      style[["numFmt"]] <- "GENERAL"
    }

    out_style <- openxlsx::createStyle(
      fontName = style[["fontName"]],
      fontSize = style[["fontSize"]],
      fontColour = font_colour_input,
      numFmt = style[["numFmt"]],
      border = NULL,
      borderColour = getOption("openxlsx.borderColour", "black"),
      borderStyle = getOption("openxlsx.borderStyle", "thin"),
      bgFill =NULL,
      fgFill = NULL,
      halign = style[["halign"]],
      valign = style[["valign"]],
      textDecoration = style[["textDecoration"]],
      wrapText = FALSE,
      textRotation = NULL,
      indent = style[["indent"]])

    if(suppressWarnings(!is.na(as.integer(style[["fontColour"]])))){
      out_style$fontColour <- c(theme = style[["fontColour"]])
    }

    return(out_style)
  } else {

    out_style <- list()

    out_style[["fontName"]] <- style$style$fontName
    out_style[["fontSize"]] <- style$style$fontSize
    out_style[["fontColour"]] <- style$style$fontColour
    ## Number formats are read differently
    # out_style[["numFmt"]] <- "GENERAL"
    # border not yet supported due to weird way it converts on style
    # out_style[["border"]] <- style$style$border
    # borderColour - not yet supported
    # borderStyle - not yet supported
    # out_style[["bgFill"]] <- style$style$bgFill
    # out_style[["fgFill"]] <- style$style$fgFill
    out_style[["halign"]] <- style$style$halign
    out_style[["valign"]] <- style$style$valign
    if(length(style$style$fontDecoration) == 0){
      out_style["textDecoration"] <- NULL
    } else {
      out_style["textDecoration"] <- style$style$fontDecoration
    }
    # wrapText - not yet supported
    # textRotation - not yet supported
    out_style["indent"] <- style$style$indent

    # Additional variables to convert
    out_style["rowHeight"] <- style$rowHeight

    return(out_style)
  }
}


add_styles_to_wb <- function(tab, add_from = c("title","headers","body")){

  # build up a table of cell style tables based on add_from
  init <- TRUE
  if("title" %in% add_from){
    if(init){
      full_table <- title_get_cell_styles_table(tab)
      init <- FALSE
    }
  }
  if("headers" %in% add_from){
    if(init){
      full_table <- top_headers_get_cell_styles_table(tab)
      init <- FALSE
    } else {
      full_table <- rbind(full_table, top_headers_get_cell_styles_table(tab))
    }
  }
  if("body" %in% add_from){
    if(init){
      full_table <- body_get_cell_styles_table(tab)
      init <- FALSE
    } else {
      full_table <- rbind(full_table, body_get_cell_styles_table(tab))
    }
  }

  # Get a unique vector of style_name from full_table
  unique_styles_definitions <- unique(full_table$style_name)

  # Add unique style_name vector to style catalogue
  tab$style_catalogue <- add_style_defintions_to_catelogue(tab$style_catalogue, unique_styles_definitions)

  # add the style_key to each style name in full table
  full_table$style_key <- unlist(tab$style_catalogue[full_table$style_name])

  # iterate over a unique list of style_keys for and apply them to the workbook for each row col
  unique_style_keys <- unique(full_table$style_key)

  for (sk in unique_style_keys){
    row_col_styles <- full_table[full_table$style_key == sk,]
    rows <- row_col_styles$row
    cols <- row_col_styles$col

    created_style <- convert_style_object(sk, convert_to_S4 = TRUE)
    openxlsx::addStyle(tab$wb, tab$wb$sheet_names[1], created_style, rows, cols)
  }

  tab
}

