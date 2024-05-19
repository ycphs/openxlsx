

#' @name write.xlsx
#' @title write data to an xlsx file
#' @description write a data.frame or list of data.frames to an xlsx file
#' @author Alexander Walker, Jordan Mark Barbone
#' @inheritParams buildWorkbook
#' @param file A file path to save the xlsx file
#' @param overwrite Overwrite existing file (Defaults to `TRUE` as with `write.table`)
#' @param ... Additional arguments passed to [buildWorkbook()]; see details
#' 
#' @inheritSection buildWorkbook Optional Parameters
#'
#' @seealso [addWorksheet()]
#' @seealso [writeData()]
#' @seealso [createStyle()] for style parameters
#' @seealso [buildWorkbook()]
#' @return A workbook object
#' @examples
#'
#' ## write to working directory
#' options("openxlsx.borderColour" = "#4F80BD") ## set default border colour
#' \dontrun{
#' write.xlsx(iris, file = "writeXLSX1.xlsx", colNames = TRUE, borders = "columns")
#' write.xlsx(iris, file = "writeXLSX2.xlsx", colNames = TRUE, borders = "surrounding")
#' }
#'
#'
#' hs <- createStyle(
#'   textDecoration = "BOLD", fontColour = "#FFFFFF", fontSize = 12,
#'   fontName = "Arial Narrow", fgFill = "#4F80BD"
#' )
#' \dontrun{
#' write.xlsx(iris,
#'   file = "writeXLSX3.xlsx",
#'   colNames = TRUE, borders = "rows", headerStyle = hs
#' )
#' }
#'
#' ## Lists elements are written to individual worksheets, using list names as sheet names if available
#' l <- list("IRIS" = iris, "MTCATS" = mtcars, matrix(runif(1000), ncol = 5))
#' \dontrun{
#' write.xlsx(l, "writeList1.xlsx", colWidths = c(NA, "auto", "auto"))
#' }
#'
#' ## different sheets can be given different parameters
#' \dontrun{
#' write.xlsx(l, "writeList2.xlsx",
#'   startCol = c(1, 2, 3), startRow = 2,
#'   asTable = c(TRUE, TRUE, FALSE), withFilter = c(TRUE, FALSE, FALSE)
#' )
#' }
#'
#' # specify column widths for multiple sheets
#' \dontrun{
#' write.xlsx(l, "writeList2.xlsx", colWidths = 20)
#' write.xlsx(l, "writeList2.xlsx", colWidths = list(100, 200, 300))
#' write.xlsx(l, "writeList2.xlsx", colWidths = list(rep(10, 5), rep(8, 11), rep(5, 5)))
#' }
#'
#' @export
write.xlsx <- function(x, file, asTable = FALSE, overwrite = TRUE, ...) {
  if ("matrix" %in% class(x)){
    x <- as.data.frame(x)
  }
  wb <- buildWorkbook(x, asTable = asTable, ...)
  saveWorkbook(wb, file = file, overwrite = overwrite)
  invisible(wb)
}
