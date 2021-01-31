#include <Rcpp.h>
#include <sstream>
#include "pugixml.hpp"

/*
 * creates output of a <row> either as <row/> if nothing else is available or as
 * <row>...</row> filling ... with <c> embedding <f> and/or <v>. If <row ...> 
 * has attributes, they are written too.
 * Known attributes for <c> are r="", s="" and t="".
 */
std::string setXMLrow(Rcpp::CharacterVector row_style, 
                      Rcpp::List cell_frm,
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
  // Rf_PrintValue(cell_frm);
  
  // Rf_PrintValue(cell_row);
  // Rf_PrintValue(cell_str);
  // Rf_PrintValue(cell_typ);

  // not if only row attribs are present
  // Rcpp::Rcout << "c_beg" << std::endl;
  if (cell_row.length() > 0)
    for (auto i = 0; i < cell_typ.length(); ++i) {
      
      // Rcpp::CharacterVector c_typ = "1"; // cell_typ[i];
      // Rcpp::CharacterVector v_typ = "2"; // cell_val[i];
      
      // Rcpp::Rcout << (i+1) << "/" << cell_typ.length() << std::endl;
      
      std::string r_typ = Rcpp::as<std::string>(cell_row[i]);
      std::string s_typ = Rcpp::as<std::string>(cell_str[i]);
      std::string t_typ = Rcpp::as<std::string>(cell_typ[i]);
      
      // Rcpp::Rcout << "b_read" << std::endl;
      Rcpp::List f_typ = cell_frm[i];
      Rcpp::List v_typ = cell_val[i];
      // Rcpp::Rcout << "e_read" << std::endl;
      
      // Rf_PrintValue(f_typ);
      // Rf_PrintValue(c_typ); // is a string already
      // Rf_PrintValue(v_typ);
      
      // each <c> needs "r" and "s"
      pugi::xml_node col = row.append_child("c");
      // save attributes only if != "" (is r is always != ""?)
      if (r_typ.compare("") != 0) col.append_attribute("r") = r_typ.c_str();
      if (s_typ.compare("") != 0) col.append_attribute("s") = s_typ.c_str();
      if (t_typ.compare("") != 0) col.append_attribute("t") = t_typ.c_str();
      
      
      /*
       * <col r="A1" s="1">
       * <f>
       * <v>
       * </col>
       * 
       */
      
      // Rf_PrintValue(f_typ);
      // check if any <f> is available
      for (auto fj = 0; fj < f_typ.length(); ++fj) {
        
        // Rcpp::Rcout << r_typ << " " << s_typ << std::endl;
        // Rcpp::Rcout << "ftyp: " << (fj +1) << "/" << f_typ.length() << std::endl;

        // Rf_PrintValue(f_typ[fj]);
        std::string ft = Rcpp::as<std::string>(f_typ[fj]);
        // Rcpp::Rcout << ft << std::endl;

        pugi::xml_document fml;
        pugi::xml_parse_result result_f = fml.load_string(ft.c_str(), pugi::parse_default | pugi::parse_fragment);
        // Rcpp::Rcout << result_f.description() << std::endl;

        pugi::xml_node fv = col.append_copy(fml.document_element());
      }
      // Rcpp::Rcout << "fin" << std::endl;
      
      // check if any <v> is available
      for (auto vj = 0; vj < v_typ.length(); ++vj) {

        std::string vt = Rcpp::as<std::string>(v_typ[vj]);

        pugi::xml_document val;
        pugi::xml_parse_result result_v = val.load_string(vt.c_str(), pugi::parse_default | pugi::parse_fragment);
        // Rcpp::Rcout << result_v.description() << std::endl;

        pugi::xml_node vv = col.append_copy(val.document_element());
      }
      
    }
    // Rcpp::Rcout << "c_end" << std::endl;
    
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