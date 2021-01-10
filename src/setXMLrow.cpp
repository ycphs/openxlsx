#include <Rcpp.h>
#include <sstream>
#include "pugixml.hpp"


std::string setXMLrow(Rcpp::CharacterVector row_attr, std::string cell_xml) {
  
  Rcpp::CharacterVector nams = row_attr.names();
  
  pugi::xml_document cell;
    
  // Rcpp::Rcout << row_attr << std::endl;
  // Rcpp::Rcout << nams << std::endl;
  // Rcpp::Rcout << cell_xml << std::endl;
  pugi::xml_parse_result result = cell.load_string(cell_xml.c_str(), 
                                                   pugi::parse_default | pugi::parse_fragment);
  
  if (!result) {
    Rcpp::stop("xml import failed", result.description());
  } 
  
  pugi::xml_document doc;
  pugi::xml_node row = doc.append_child("row");
  
  for (auto i = 0; i < nams.length(); ++i) {
    row.append_attribute(nams[i]) = Rcpp::as<std::string>(row_attr[i]).c_str();
  }
  for (pugi::xml_node cld = cell.first_child();
       cld;
       cld = cld.next_sibling())
  {
    row.append_copy(cld);
  }
  
  std::ostringstream oss;
  doc.print(oss, " ");
  
  return oss.str();
}