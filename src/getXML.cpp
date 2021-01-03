#include <Rcpp.h>
#include <sstream>
#include "pugixml.hpp"

// [[Rcpp::export]]
SEXP getXML1(std::string str, std::string child) {

  pugi::xml_document doc;
  pugi::xml_parse_result result = doc.load_string(str.c_str());
  if (!result) {
    Rcpp::stop("xml import unsuccessfull");
  }


  std::vector<std::string> res;

  for (pugi::xml_node worksheet = doc.child(child.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(child.c_str()))
  {
    std::ostringstream oss;
    worksheet.print(oss, " ", pugi::format_raw);
    res.push_back(oss.str());
  }

  return  Rcpp::wrap(res);
}

// [[Rcpp::export]]
SEXP getXML2(std::string str, std::string level1, std::string child) {
  
  pugi::xml_document doc;
  pugi::xml_parse_result result = doc.load_string(str.c_str());
  if (!result) {
    Rcpp::stop("xml import unsuccessfull");
  }
  
  
  std::vector<std::string> res;
  
  for (pugi::xml_node worksheet = doc.child(level1.c_str()).child(child.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(child.c_str()))
  {
    std::ostringstream oss;
    worksheet.print(oss, " ", pugi::format_raw);
    res.push_back(oss.str());
  }
  
  return  Rcpp::wrap(res);
}

// [[Rcpp::export]]
SEXP getXML3(std::string str, std::string level1, std::string level2, std::string child) {
  
  pugi::xml_document doc;
  pugi::xml_parse_result result = doc.load_string(str.c_str());
  if (!result) {
    Rcpp::stop("xml import unsuccessfull");
  }
  
  
  std::vector<std::string> res;
  
  for (pugi::xml_node worksheet = doc.child(level1.c_str()).child(level2.c_str()).child(child.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(child.c_str()))
  {
    std::ostringstream oss;
    worksheet.print(oss, " ", pugi::format_raw);
    res.push_back(oss.str());
  }
  
  return  Rcpp::wrap(res);
}
