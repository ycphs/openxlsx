#' @method print pugi_xml
#' @param x somthing to print
#' @param raw print as raw text
#' @param ... to please check
#' @export
print.pugi_xml <- function(x, raw = FALSE, ...) {
  cat(printXPtr(x, raw))
}