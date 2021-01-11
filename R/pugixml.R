#' @method print pugi_xml
#' @export
print.pugi_xml <- function(x) {
  print(printXPtr(x))
}