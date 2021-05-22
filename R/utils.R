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
  if (is.null(x)) {
    FALSE
  } else {
    !inherits(x, class)
  }
}

is_true_false <- function(x) {
  is.logical(x) && length(x) == 1L && is.na(x)
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
# returns the current options to decrease line use
get_set_options <- function() {
  op <- options()
  options(
    # increase scipen to avoid writing in scientific
    scipen = 200,
    OutDec = ".",
    digits = 22
  )
  op
}