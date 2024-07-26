#' Build Workbook
#'
#' Build a workbook from a data.frame or named list
#'
#' @details
#' This function can be used as shortcut to create a workbook object from a
#'   data.frame or named list.  If names are available in the list they will be
#'   used as the worksheet names.  The parameters in `...` are collected
#'   and passed to [writeData()] or [writeDataTable()] to
#'   initially create the Workbook objects then appropriate parameters are
#'   passed to [setColWidths()].
#'
#' @param x A data.frame or a (named) list of objects that can be handled by
#'   [writeData()] or [writeDataTable()] to write to file
#' @param asTable If `TRUE` will use [writeDataTable()] rather
#'   than [writeData()] to write `x` to the file (default:
#'   `FALSE`)
#' @param ... Additional arguments passed to [writeData()],
#'   [writeDataTable()], [setColWidths()] (see Optional
#'   Parameters)
#' @author Jordan Mark Barbone
#' @returns A Workbook object
#'
#' @details
#' columns of x with class Date or POSIXt are automatically
#' styled as dates and datetimes respectively.
#'
#' @section Optional Parameters:
#'
#' **createWorkbook Parameters**
#' \describe{
#'   \item{**creator**}{ A string specifying the workbook author}
#' }
#'
#' **addWorksheet Parameters**
#' \describe{
#'   \item{**sheetName**}{ Name of the worksheet}
#'   \item{**gridLines**}{ A logical. If `FALSE`, the worksheet grid lines will be hidden.}
#'   \item{**tabColour**}{ Colour of the worksheet tab. A valid colour (belonging to colours())
#'   or a valid hex colour beginning with "#".}
#'   \item{**zoom**}{ A numeric between 10 and 400. Worksheet zoom level as a percentage.}
#' }
#'
#' **writeData/writeDataTable Parameters**
#' \describe{
#'   \item{**startCol**}{ A vector specifying the starting column(s) to write df}
#'   \item{**startRow**}{ A vector specifying the starting row(s) to write df}
#'   \item{**xy**}{ An alternative to specifying startCol and startRow individually.
#'  A vector of the form c(startCol, startRow)}
#'   \item{**colNames or col.names**}{ If `TRUE`, column names of x are written.}
#'   \item{**rowNames or row.names**}{ If `TRUE`, row names of x are written.}
#'   \item{**headerStyle**}{ Custom style to apply to column names.}
#'   \item{**borders**}{ Either "surrounding", "columns" or "rows" or NULL.  If "surrounding", a border is drawn around the
#' data.  If "rows", a surrounding border is drawn a border around each row. If "columns", a surrounding border is drawn with a border
#' between each column.  If "`all`" all cell borders are drawn.}
#'   \item{**borderColour**}{ Colour of cell border}
#'   \item{**borderStyle**}{ Border line style.}
#'   \item{**keepNA**}{If `TRUE`, NA values are converted to #N/A (or `na.string`, if not NULL) in Excel, else NA cells will be empty. Defaults to FALSE.}
#'   \item{**na.string**}{If not NULL, and if `keepNA` is `TRUE`, NA values are converted to this string in Excel. Defaults to NULL.}
#' }
#'
#' **freezePane Parameters**
#' \describe{
#'   \item{**firstActiveRow**}{Top row of active region to freeze pane.}
#'   \item{**firstActiveCol**}{Furthest left column of active region to freeze pane.}
#'   \item{**firstRow**}{If `TRUE`, freezes the first row (equivalent to firstActiveRow = 2)}
#'   \item{**firstCol**}{If `TRUE`, freezes the first column (equivalent to firstActiveCol = 2)}
#' }
#'
#' **colWidths Parameters**
#' \describe{
#'   \item{**colWidths**}{May be a single value for all columns (or "auto"), or a list of vectors that will be recycled for each sheet (see examples)}
#' }
#'
#' @examples
#' x <- data.frame(a = 1, b = 2)
#' wb <- buildWorkbook(x)
#'
#' y <- list(a = x, b = x, c = x)
#' buildWorkbook(y, asTable = TRUE)
#' buildWorkbook(y, asTable = TRUE, tableStyle = "TableStyleLight8")
#'
#' @seealso [write.xlsx()]
#'
#' @export

buildWorkbook <- function(x, asTable = FALSE, ...) {
  if (!is.logical(asTable)) {
    stop("asTable must be a logical.")
  }

  params <- list(...)
  isList <- inherits(x, "list")

  if (isList) {
    params[["sheetName"]] <- params[["sheetName"]] %||% names(x) %||% paste0("Sheet ", seq_along(x))
  }

  ## create new Workbook object
  wb <- do_call_params(createWorkbook, params)

  ## If a list is supplied write to individual worksheets using names if available
  if (isList) {
    do_call_params(addWorksheet, params, wb = list(wb), .map = TRUE)
  } else {
    params[["sheetName"]] <- params[["sheetName"]] %||% "Sheet 1"
    do_call_params(addWorksheet, params, wb = wb)
  }

  params[["sheet"]] <- params[["sheet"]] %||% params[["sheetName"]]

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

  params[["startCol"]] <- params[["startCol"]] %||% 1
  params[["startCol"]] <- rep_len(list(params[["startCol"]]), length.out = length(x))
  params[["colWidths"]] <- params[["colWidths"]] %||% ""
  params[["colWidths"]] <- rep_len(as.list(params[["colWidths"]]), length.out = length(x))

  for (i in seq_along(wb[["worksheets"]])) {
    if (identical(params[["colWidths"]][[i]], "auto")) {
      setColWidths(
        wb,
        sheet = i,
        cols = seq_along(x[[i]]) + params[["startCol"]][[i]] - 1L,
        widths = "auto"
      )
    } else if (!identical(params[["colWidths"]][[i]], "")) {
      setColWidths(
        wb,
        sheet = i,
        cols = seq_along(x[[i]]) + params[["startCol"]][[i]] - 1L,
        widths = params[["colWidths"]][[i]]
      )
    }
  }
  wb
}
