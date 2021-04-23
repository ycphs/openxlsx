.onAttach <- function(libname, pkgname) {
  op <- options()
  toset <- !(names(op.openxlsx) %in% names(op))
  if (any(toset)) {
    options(op.openxlsx[toset]) 
  }
}
