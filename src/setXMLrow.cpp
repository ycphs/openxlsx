#include <Rcpp.h>
#include <sstream>
#include "pugixml.hpp"


std::string setXMLrow(Rcpp::CharacterVector row_style, Rcpp::List cell_typ, Rcpp::List cell_val) {
  
  Rcpp::CharacterVector nams = row_style.names();
  pugi::xml_document typ;
  pugi::xml_document val;
  
  
  pugi::xml_document doc;
  pugi::xml_node row = doc.append_child("row");
  
  for (auto i = 0; i < nams.length(); ++i) {
    row.append_attribute(nams[i]) = Rcpp::as<std::string>(row_style[i]).c_str();
  }
  
  // Rcpp::Rcout << row_attr << std::endl;
  // Rcpp::Rcout << nams << std::endl;
  // Rcpp::Rcout << cell_xml << std::endl;
  
  for (auto i = 0; i < cell_typ.length(); ++i) {
    
    Rcpp::CharacterVector c_typ = "1"; // cell_typ[i];
    Rcpp::CharacterVector v_typ = "2"; // cell_val[i];
    
    Rcpp::CharacterVector tnams = ""; //c_typ.names();
    
    
    pugi::xml_node col = row.append_child("c");
    
    for (auto j = 0; j < c_typ.length(); ++j) {
      col.append_attribute(tnams[j]) = Rcpp::as<std::string>(c_typ[j]).c_str();
    }
    
    
    for (pugi::xml_node ti = col.first_child();
         ti;
         ti = ti.next_sibling())
    {
      
      col.append_copy(ti);
      
      if (v_typ.length() > 0) {

        pugi::xml_parse_result result2 = val.load_string(Rcpp::as<std::string>(v_typ).c_str(),
                                                         pugi::parse_default | pugi::parse_fragment);
        pugi::xml_node val = col.append_child("v");

        for (pugi::xml_node vi = val.first_child();
             vi;
             vi = vi.next_sibling())
        {
          col.append_copy(vi);
        }
      }
      
    }
  }
  
  std::ostringstream oss;
  doc.print(oss, " ", pugi::format_raw);
  
  return oss.str();
}