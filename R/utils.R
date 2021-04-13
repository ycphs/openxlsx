
`%||%` <- function(x, y) if (is.null(x)) y else x

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


foo <- function(a, b) a + b
foo(1 , 2)
bar <- function(...) foo(...)
bar(a = 1, b = 2)
do.call(bar, list(a = 1, b = 2))

g <- function(...) mapply(bar, ...)
do.call(g, list(a = 1:2, b = 2))

