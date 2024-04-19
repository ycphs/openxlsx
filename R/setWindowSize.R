#' Set and Get Window Size for xlsx file
#'
#' @param wb A Workbook object
#' @param xWindow the horizontal coordinate of the top left corner of the window
#' @param yWindow the vertical coordinate of the top left corner of the window
#' @param windowWidth the width of the window
#' @param windowHeight the height of the window
#' 
#' Set the size and position of the window when you open the xlsx file.  The units are in twips.  See
#' [Microsoft's documentation for the xlsx standard](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.workbookview?view=openxml-2.8.1)
#'
#' @export
#'
#' @examples
#' ## Create Workbook object and add worksheets
#' wb <- createWorkbook()
#' addWorksheet(wb, "S1")
#' getWindowSize(wb)
#' setWindowSize(wb, windowWidth = 10000)
setWindowSize <- function(wb, xWindow = NULL, yWindow = NULL, windowWidth = NULL, windowHeight= NULL) {

  bookViews <- wb$workbook$bookViews

  if(!is.null(xWindow)) {
    if(as.integer(xWindow) >= 0L) {
      bookViews <- sub("xWindow=\"\\d+", paste0("xWindow=\"", xWindow), bookViews)
    } else {
      stop("xWindow must be >= 0")
    }
  }

  if(!is.null(yWindow)) {
    if(as.integer(yWindow) >= 0L) {
      bookViews <- sub("yWindow=\"\\d+", paste0("yWindow=\"", yWindow), bookViews)
    } else {
      stop("yWindow must be >= 0")
    }
  }

  if(!is.null(windowWidth)) {
    if(as.integer(windowWidth) >= 100L) {
      bookViews <- sub("windowWidth=\"\\d+", paste0("windowWidth=\"", windowWidth), bookViews)
    } else {
      stop("windowWidth must be >= 100")
    }
  }

  if(!is.null(windowHeight)) {
    if(as.integer(windowHeight) >= 100L) {
      bookViews <- sub("windowHeight=\"\\d+", paste0("windowHeight=\"", windowHeight), bookViews)
    } else {
      stop("windowHeight must be >= 100")
    }
  }

  wb$workbook$bookViews <- bookViews
}

#' @rdname setWindowSize
#' @export

getWindowSize <- function(wb) {
  bookViews <- wb$workbook$bookViews

  c(getAttrs(bookViews, "xWindow"),
    getAttrs(bookViews, "yWindow"),
    getAttrs(bookViews, "windowWidth"),
    getAttrs(bookViews, "windowHeight"))

}
