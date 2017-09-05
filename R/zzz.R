xltabr_default_options <- list(
  xltabr.style.path = system.file("extdata", "styles.xlsx", package = "xltabr"),
  xltabr.cell.format.path = system.file("extdata", "style_to_excel_number_format.csv", package = "xltabr"),
  xltabr.number.format.path = system.file("extdata", "number_format_defaults.csv", package = "xltabr"),
  xltabr.style.override.path = NULL)

.onLoad <- function(libname, pkgname) {
  op <- options()
  toset <- !(names(xltabr_default_options) %in% names(op))
  if(any(toset)) options(xltabr_default_options[toset])

  invisible()
}

.onUnload <- function(libname, pkgname){
  options(list(xltabr.style.path = NULL,
               xltabr.cell.format.path = NULL,
               xltabr.number.format.path = NULL,
               xltabr.style.override.path = NULL))
}
