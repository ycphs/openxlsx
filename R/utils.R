#' If NULL then ...
#'
#' Replace NULL
#'
#' @param x A value to check
#' @param y A value to substitute if x is null
#' @examples
#' \dontrun{
#' x <- NULL
#' x <- x %||% "none"
#' x <- x %||% NA
#' }
#'
#' @name if_null_then
`%||%` <- function(x, y) if (is.null(x)) y else x

is_not_class <- function(x, class) {
  !(inherits(x, class) | is.null(x))
}

is_true_false <- function(x) {
  is.logical(x) && length(x) == 1L && !is.na(x)
}

pair_rc <- function(r, c) {
  # Assumes that there can't be more than 10M columns
  # Excel allows 16,384
  # Google Sheets allows 18,278
  r+c*1e-7
}

do_call_params <- function(fun, params, ..., .map = FALSE) {
  fun <- match.fun(fun)
  call_params <- c(list(...), params[names(params) %in% names(formals(fun))])
  call_params <- lapply(call_params, function(x) if (is.object(x)) list(x) else x)

  call_fun <- if (.map) {
    function(...) mapply(fun, ..., MoreArgs = NULL, SIMPLIFY = FALSE, USE.NAMES = FALSE)
  } else {
    fun
  }

  do.call(call_fun, call_params)
}

# sets temporary options
# option() returns the original values
get_set_options <- function() {
  options(
    # increase scipen to avoid writing in scientific
    scipen = 200,
    OutDec = ".",
    digits = 22
  )
}


#' helper function to create tempory directory for testing purpose
#' @param name for the temp file
#' @export
temp_xlsx <- function(name = "temp_xlsx") {
  tempfile(pattern = paste0(name, "_"), fileext = ".xlsx")
}
