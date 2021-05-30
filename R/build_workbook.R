#' Build Workbook
#' 
#' Build a workbook from a data.frame or named list
#' 
#' @details 
#' This function can be used as shortcut to create a workbook object from a 
#'   data.frame or named list.  If names are available in the list they will be
#'   used as the worksheet names.  The parameters in \code{...} are collected
#'   and passed to \code{\link{writeData}} or \code{\link{writeDataTable}} to
#'   initially create the Workbook objects then appropriate parameters are 
#'   passed to \code{\link{setColWidths}}.
#' 
#' @param x A data.frame or a (named) list of objects that can be handled by
#'   \code{\link{writeData}} or \code{\link{writeDataTable}} to write to file
#' @param asTable If \code{TRUE} will use \code{\link{writeDataTable}} rather
#'   than \code{\link{writeData}} to write \code{x} to the file (default: 
#'   \code{FALSE})
#' @param ... Additional arguments passed to \code{\link{writeData}}, 
#'   \code{\link{writeDataTable}}, \code{\link{setColWidths}}
#' @author Jordan Mark Barbone
#' @returns A Workbook object
#' 
#' @examples
#' x <- data.frame(a = 1, b = 2)
#' wb <- buildWorkbook(x)
#' 
#' y <- list(a = x, b = x, c = x)
#' buildWorkbook(y, asTable = TRUE)
#' buildWorkbook(y, asTable = TRUE, tableStyle = "TableStyleLight8")
#' 
#' @seealso \code{\link{write.xlsx}}
#' 
#' @export

buildWorkbook <- function(x, asTable = FALSE, ...) {
  if (!is.logical(asTable)) {
    stop("asTable must be a logical.")
  }
  
  params <- list(...)
  isList <- inherits(x, "list")
  
  if (isList) {
    params$sheetName <- params$sheetName %||% names(x) %||% paste0("Sheet ", seq_along(x))
  }
  
  ## create new Workbook object
  wb <- do_call_params(createWorkbook, params)
  
  ## If a list is supplied write to individual worksheets using names if available
  if (isList) {
    do_call_params(addWorksheet, params, wb = list(wb), .map = TRUE)
  } else {
    params$sheetName <- params$sheetName %||% "Sheet 1"
    do_call_params(addWorksheet, params, wb = wb)
  }
  
  params$sheet <- params$sheet %||% params$sheetName
  
  # write Data
  if (asTable) {
    do_call_params(writeDataTable, params, x = x, wb = list(wb), .map = TRUE)
  } else {
    do_call_params(writeData, params, x = x, wb = wb, .map = TRUE)
  }
  
  do_setColWidths(wb, x, params, isList)
  do_call_params(freezePane, params, wb = list(wb), .map = TRUE)
  wb
}


do_setColWidths <- function(wb, x, params, isList) {
  if (!isList) {
    x <- list(x)
  }
  
  params$startCol <- params$startCol %||% 1
  params$startCol <- rep_len(list(params$startCol), length.out = length(x))
  params$colWidths <- params$colWidths %||% ""
  params$colWidths <- rep_len(as.list(params$colWidths), length.out = length(x))
  
  for (i in seq_along(wb$worksheets)) {
    if (identical(params$colWidths[[i]], "auto")) {
      setColWidths(
        wb,
        sheet = i,
        cols = seq_along(x[[i]]) + params$startCol[[i]] - 1L, 
        widths = "auto"
      )
    } else if (!identical(params$colWidths[[i]], "")) {
      setColWidths(
        wb,
        sheet = i,
        cols = seq_along(x[[i]]) + params$startCol[[i]] - 1L,
        widths = params$colWidths[[i]]
      )
    }
  }
  wb
}
