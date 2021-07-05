#' xlsx reading, writing and editing.
#'
#' openxlsx simplifies the the process of writing and styling Excel xlsx files from R
#' and removes the dependency on Java.
#'
#' @name openxlsx
#' @docType package
#' @useDynLib openxlsx, .registration=TRUE
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
#'    \item{\code{\link{op.openxlsx}}}
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
#'  \item{openxlsx.creator}{""}
#'  \item{openxlsx.dateFormat}{"mm/dd/yyyy"}
#'  \item{openxlsx.datetimeFormat}{"yyyy-mm-dd hh:mm:ss"}
#'  \item{openxlsx.headerStyle}{NULL}
#'  \item{openxlsx.keepNA}{FALSE}
#'  \item{openxlsx.na.string}{NULL}
#'  \item{openxlsx.numFmt}{NULL}
#'  \item{openxlsx.orientation}{"portrait"}
#'  \item{openxlsx.paperSize}{9}
#'  \item{openxlsx.tabColour}{"TableStyleLight9"}
#'  \item{openxlsx.tableStyle}{"TableStyleLight9"}
#'  \item{openxlsx.withFilter}{NA Whether to write data with or without a 
#'  filter. If NA will make filters with \code{writeDataTable} and will not for
#'  \code{writeData}}
#' }
#' 
#' See the Formatting vignette for examples.
#'
#' Additional options
#'
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
#'   default option from \code{op.openxlsx}.  This will also check that the
#'   intended option is a standard option (listed in \code{op.openxlsx}) and
#'    will provide a warning otherwise.
#'   
#' \code{openxlsx_setOp()} is a safer way to set an option as it will first
#'   check that the option is a standard option (as above) before setting.
#' 
#' @examples
#' openxlsx_getOp("borders")
#' op.openxlsx[["openxlsx.borders"]]
#' 
#' @export
#' @name openxlsx_options
op.openxlsx <- list(
  openxlsx.bandedCols       = FALSE,
  openxlsx.bandedRows       = TRUE,
  openxlsx.borderColour     = "black",
  openxlsx.borders          = NULL,
  openxlsx.borderStyle      = "thin",
  # Where is compressionLevel called?
  openxlsx.compressionLevel = 9,
  openxlsx.creator          = "",
  openxlsx.dateFormat       = "mm/dd/yyyy",
  openxlsx.datetimeFormat   = "yyyy-mm-dd hh:mm:ss",
  openxlsx.hdpi             = 300,
  openxlsx.header           = NULL,
  openxlsx.headerStyle      = NULL,
  openxlsx.firstColumn      = NULL,
  openxlsx.firstFooter      = NULL,
  openxlsx.firstHeader      = NULL,
  openxlsx.footer           = NULL,
  openxlsx.evenFooter       = NULL,
  openxlsx.evenHeader       = NULL,
  openxlsx.gridLines        = TRUE,
  openxlsx.keepNA           = FALSE,
  openxlsx.lastColumn       = NULL,
  openxlsx.na.string        = NULL,
  openxlsx.maxWidth         = 250,
  openxlsx.minWidth         = 3,
  openxlsx.numFmt           = "GENERAL",
  openxlsx.oddFooter        = NULL,
  openxlsx.oddHeader        = NULL,
  openxlsx.orientation      = "portrait",
  openxlsx.paperSize        = 9,
  openxlsx.showGridLines    = NA,
  openxlsx.tabColour        = NULL,
  openxlsx.tableStyle       = "TableStyleLight9",
  openxlsx.vdpi             = 300,
  openxlsx.withFilter       = NULL
)


#' @param x An option name (\code{"openxlsx."} prefix optional)
#' @param default A default value if \code{NULL}
#' @rdname openxlsx_options
#' @export
openxlsx_getOp <- function(x, default = NULL) {
  if (length(x) != 1L || length(default) > 1L) {
   stop("x must be length 1 and default NULL or length 1", call. = FALSE)
  } 
  
  x <- check_openxlsx_op(x)
  getOption(x, op.openxlsx[[x]]) %||% default
}

#' @param value The new value for the option (optional if x is a named list)
#' @rdname openxlsx_options
#' @export
openxlsx_setOp <- function(x, value) {
  if (is.list(x)) {
    if (is.null(names(x))) {
      stop("x cannot be an unnamed list", call. = FALSE)
    }
    
    mapply(openxlsx_setOp, x = names(x), value = x)
  }
  
  value <- as.list(value)
  names(value) <- check_openxlsx_op(x)
  options(value)
}

check_openxlsx_op <- function(x) {
  if (length(x) != 1L || !is.character(x)) {
    stop("option must be a character vector of length 1", call. = FALSE)
  }
  
  if (!grepl("^openxlsx[.]", x)) {
    x <- paste0("openxlsx.", x)
  }
  
  if (!x %in% names(op.openxlsx)) {
    warning(
      x, " is not a standard openxlsx option\nCheck spelling",
      call. = FALSE
    )
  }
  
  x
}
