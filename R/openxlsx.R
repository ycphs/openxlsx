#' xlsx reading, writing and editing.
#'
#' openxlsx simplifies the the process of writing and styling Excel xlsx files from R
#' and removes the dependency on Java.
#'
#' @name openxlsx
#' @docType package
#' @useDynLib openxlsx, .registration=TRUE
#' @importFrom Rcpp sourceCpp
#' @importFrom zip zipr
#' @importFrom utils download.file head menu unzip
#'
#' @seealso
#' \itemize{
#'    \item{\code{vignette("Introduction", package = "openxlsx")}}
#'    \item{\code{vignette("formatting", package = "openxlsx")}}
#'    \item{\code{\link{writeData}}}
#'    \item{\code{\link{writeDataTable}}}
#'    \item{\code{\link{write.xlsx}}}
#'    \item{\code{\link{read.xlsx}}}
#'    \item{\code{\link{op.openxlx}}}
#'   }
#' for examples
#'
#' @details
#' The openxlsx package uses global options, most to simplify formatting.  These
#'   are stored in the \code{op.openxlsx} object.
#'
#' \describe{
#'  \item{openxlsx.bandedCols}{FALSE}
#'  \item{openxlsx.bandedRows}{TRUE}
#'  \item{openxlsx.borderColour}{"black"}
#'  \item{openxlsx.borders}{"none"}
#'  \item{openxlsx.borderStyle}{"thin"}
#'  \item{openxlsx.compressionLevel}{"9"}
#'  \item{openxlsx.dateFormat}{"mm/dd/yyyy"}
#'  \item{openxlsx.datetimeFormat}{"yyyy-mm-dd hh:mm:ss"}
#'  \item{openxlsx.headerStyle}{NULL}
#'  \item{openxlsx.na.string}{NULL}
#'  \item{openxlsx.numFmt}{NULL}
#'  \item{openxlsx.orientation}{"portrait"}
#'  \item{openxlsx.paperSize}{9}
#'  \item{openxlsx.tabColour}{"TableStyleLight9"}
#'  \item{openxlsx.tableStyle}{"TableStyleLight9"}
#'  \item{openxlsx.withFilter}{FALSE}
#' }
#' 
#' See the Formatting vignette for examples.
#'
#' Additional options
#'
#' \describe{
#' }
#'
NULL


#' openxlsx Options
#' 
#' See and get the openxlsx options
#' 
#' @details
#' 
#' \code{openxlsx_getOp()} retrieves the \code{"openxlsx"} options found in
#'   \code{op.openxlsx}.  If none are set (currently `NULL`) retrieves the 
#'   default option from \code{op.openxlsx}
#' 
#' @param x An option name (\code{"openxlsx."} prefix optional)
#' 
#' @examples
#' openxlsx_getOp("borders")
#' op.openxlsx[["openxlsx.borders"]]
#' 
#' @export
#' @name openxlsx_options
openxlsx_getOp <- function(x) {
  if (length(x) != 1L || !is.character(x)) {
    stop("option must be a character vector of length 1", call. = FALSE)
  }
  
  ind <- grep("^openxlsx[.]", x, invert = TRUE)
  x[ind] <- paste0("openxlsx.", x[ind])
  check_openxlsx_op(x)
}

check_openxlsx_op <- function(x) {
  if (!x %in% names(op.openxlsx)) {
    warning(
      x, " is not a standard openxlsx option\nCheck spelling",
      call. = FALSE
    )
  }
  getOption(x, op.openxlsx[[x]])
}

#' @rdname openxlsx_options
#' @export
op.openxlsx <- list(
  openxlsx.bandedCols = FALSE,
  openxlsx.bandedRows = TRUE,
  openxlsx.borderColour = "black",
  openxlsx.borders = "none",
  openxlsx.borderStyle = "thin",
  openxlsx.compressionLevel = "9",
  openxlsx.dateFormat = "mm/dd/yyyy",
  openxlsx.datetimeFormat = "yyyy-mm-dd hh:mm:ss",
  openxlsx.headerStyle = NULL,
  openxlsx.na.string = NULL,
  # TODO jmb Should numFmt be something other than NULL?
  openxlsx.numFmt = NULL,
  openxlsx.orientation = "portrait",
  openxlsx.paperSize = 9,
  openxlsx.tabColour = "TableStyleLight9",
  openxlsx.tableStyle = "TableStyleLight9",
  openxlsx.withFilter = FALSE
)

.onAttach <- function(libname, pkgname) {
  # Set options if NULL
  toset <- !names(op.openxlsx) %in% names(options())
  if (any(toset)) {
    options(op.openxlsx[toset]) 
  }
}
