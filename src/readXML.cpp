#include <Rcpp.h>
#include <sstream>
#include "pugixml.hpp"

// [[Rcpp::export]]
SEXP readXML(std::string path) {
  
  pugi::xml_document doc;
  pugi::xml_parse_result result = doc.load_file(path.c_str());
  if (!result) {
    Rcpp::stop("xml import unsuccessfull");
  }
  
  std::ostringstream oss;
  doc.print(oss, " ", pugi::format_raw);
  
  return  Rcpp::wrap(oss.str());
}