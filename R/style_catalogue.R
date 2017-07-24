# Use default .xlsx style catalogue to initialise the catalogue of styles
style_catalogue_initialise <- function (tab) {
  path <- system.file("extdata", "styles.xlsx", package = "xltabr")
  tab <- style_catalogue_xlsx_import(tab, path)
  tab
}

# Iterate through the default style workbook, adding them to the styles catalogue
style_catalogue_xlsx_import <- function(tab, path) {

  wb <- openxlsx::loadWorkbook(path)
  openxlsx::readWorkbook(wb)

  for (i in wb$styleObjects) {
    r <- i$rows
    c <- i$cols
    cell <- openxlsx::readWorkbook(wb, rows = r, cols = c, colNames = FALSE, rowNames = FALSE)

    value <- cell[1, 1]

    tmp_list <- list()
    tmp_list$style <- i$style

    cell <- openxlsx::readWorkbook(wb, rows = r, cols = c + 1, colNames = FALSE, rowNames = FALSE)
    tmp_list$row_height <- cell[1, 1]

    style_key <- create_style_key(tmp_list)

    # If the style_key does not exist in the style_dictionary values then S4 object to style catalogue
    if (!style_key %in% unlist(tab$style_dictionary)){
      tab$style_catalogue[[style_key]] <- tmp_list
    }
    # Always log the key value pair in the style dictionary
    tab$style_dictionary[[value]] <- style_key
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
create_style_key <- function(sc){

  # NULL returns from property_to_key are automatically removed from vector
  style_properties <- c(property_to_key(sc$style, "fontName"),
                        property_to_key(sc$style, "fontSize"),
                        property_to_key(sc$style, "fontColour"),
                        property_to_key(sc$style, "numFmt"),
                        property_to_key(sc$style, "border"),
                        property_to_key(sc$style, "borderColour"),
                        property_to_key(sc$style, "bordersc$style"),
                        property_to_key(sc$style, "bgFill"),
                        property_to_key(sc$style, "fgFill"),
                        property_to_key(sc$style, "halign"),
                        property_to_key(sc$style, "valign"),
                        property_to_key(sc$style, "fontDecoration"),
                        property_to_key(sc$style, "wrapText"),
                        property_to_key(sc$style, "textRotation"),
                        property_to_key(sc$style, "indent"),
                        paste0("rowHeight_", sc$row_height))

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
style_key_passer <- function(style_key){
  specs <- unlist(strsplit(style_key, "\\|"))

  # utilising Rs stupid ordered hashing for the first and probably last time.
  # Can interate through style categories in order to get inherited changes from left to right.
  style_categories <- list()
  for (i in 1:length(specs)){
    seperate_property <- unlist(strsplit(specs[[i]], "_"))
    style_categories[seperate_property[1]] <- unlist(strsplit(seperate_property[2], "%"))
  }

  style_categories
}

# looks at the cell_style_definition and build a final style_key for that cell
build_style <- function(cell_style_definition){
  # Convert cell style inheritence string into an array
  seperated_style_definition <- unlist(strsplit(cell_style_definition, "\\|"))
  if (length(seperated_style_definition) <= 1){
    # If only one style (e.g body then return it's style_key)
    return (style_dictionary[[style]])
  }
  else{
    # Otherwise build final style
    previous_style <- style_key_passer(style_dictionary[[seperated_style_definition[1]]])
    for (i in 2:length(seperated_style_definition)){
      current_style <- style_key_passer(style_dictionary[[seperated_style_definition[i]]])
      for (property_name in names(current_style)){
        previous_style[property_name] <- current_style[property_name]
      }
    }
  }
  return (create_style_key(previous_style))
}

