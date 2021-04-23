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
