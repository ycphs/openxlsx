#include <Rcpp.h>
#include <sstream>
#include "pugixml.hpp"


std::string setXMLrow(Rcpp::CharacterVector row_style, 
                      Rcpp::List cell_typ,
                      Rcpp::List cell_val,
                      Rcpp::List cell_row,
                      Rcpp::List cell_str) {
  
  Rcpp::CharacterVector nams = row_style.names();
  pugi::xml_document typ;
  
  
  pugi::xml_document doc;
  pugi::xml_node row = doc.append_child("row");
  
  for (auto i = 0; i < nams.length(); ++i) {
    row.append_attribute(nams[i]) = Rcpp::as<std::string>(row_style[i]).c_str();
  }
  
  // Rcpp::Rcout << row_attr << std::endl;
  // Rcpp::Rcout << nams << std::endl;
  // Rcpp::Rcout << cell_xml << std::endl;
  
  // Rcpp::Rprintf(cell_typ);
  
  // not if only row attribs are present
  if (cell_row.length() > 0)
    for (auto i = 0; i < cell_typ.length(); ++i) {
      
      // Rcpp::CharacterVector c_typ = "1"; // cell_typ[i];
      // Rcpp::CharacterVector v_typ = "2"; // cell_val[i];
      
      
      std::string r_typ = Rcpp::as<std::string>(cell_row[i]);
      std::string s_typ = Rcpp::as<std::string>(cell_str[i]);
      Rcpp::List c_typ = cell_typ[i];
      Rcpp::List v_typ = cell_val[i];
      
      // Rf_PrintValue(c_typ);
      // Rf_PrintValue(v_typ);
      
      pugi::xml_node col = row.append_child("c");
      col.append_attribute("r") = r_typ.c_str();
      col.append_attribute("s") = s_typ.c_str();
      
      for (auto j = 0; j < c_typ.length(); ++j) {
        Rcpp::CharacterVector celltyp = c_typ[j];
        col.append_attribute("t") = celltyp;
      }
      
      // std::string rt = Rcpp::as<std::string>(r_typ[j]);
      // std::string st = Rcpp::as<std::string>(s_typ[j]);
      
      for (auto j = 0; j < v_typ.length(); ++j) {
        
        std::string vt = Rcpp::as<std::string>(v_typ[j]);
        
        pugi::xml_document val;
        pugi::xml_parse_result result2 = val.load_string(vt.c_str(), pugi::parse_default | pugi::parse_fragment);
        // Rcpp::Rcout << result2.description() << std::endl;
        
        pugi::xml_node vv = col.append_copy(val.document_element());
        
      }
    }
    
    std::ostringstream oss;
  // doc.print(oss, " ");
  doc.print(oss, " ", pugi::format_raw);
  
  return oss.str();
}

// [[Rcpp::export]]
SEXP setXMLcols(Rcpp::List cols_attr) {
  
  Rcpp::CharacterVector z(cols_attr.length());
  
  for (auto i = 0; i < cols_attr.length(); ++i) {
    
    pugi::xml_document doc;
    pugi::xml_node row = doc.append_child("col");
    
    Rcpp::CharacterVector col = cols_attr[i];
    
    for (auto j = 0; j < col.length(); ++j) {
      Rcpp::CharacterVector attrnams = col.names();
      row.append_attribute(attrnams[j]) = Rcpp::as<std::string>(col[j]).c_str();
    }
    
    std::ostringstream oss;
    doc.print(oss, " ", pugi::format_raw);
    
    z[i] = oss.str();
  }
  
  return z;
}