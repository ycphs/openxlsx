#include "openxlsx_types.h"

// [[Rcpp::export]]
std::string set_row(Rcpp::List row_attr, Rcpp::List col_vals, Rcpp::List col_attr) {
  
  pugi::xml_document doc;
  
  pugi::xml_node row = doc.append_child("row");
  
  for (auto i = 0; i < row_attr.length(); ++i) {
    Rcpp::CharacterVector attrnams = row_attr.names();
    row.append_attribute(attrnams[i]) = Rcpp::as<std::string>(row_attr[i]).c_str();
  }
  
  for (auto i = 0; i < col_attr.length(); ++i) {
    
    // create node <c>
    pugi::xml_node cell = row.append_child("c");

    Rcpp::List cell_atr = col_attr[i];
    Rcpp::List cell_val = col_vals[i];
    
    std::vector<std::string> cell_atr_names = cell_atr.names();
    std::vector<std::string> cell_val_names = cell_val.names();
    
    // append attributes <c r="A1" ...>
    for (auto j = 0; j < cell_atr.length(); ++j) {
      std::string c = cell_atr[j];
      cell.append_attribute(cell_atr_names[j].c_str()) = c.c_str();
    }
    
    // append nodes <c r="A1" ...><v>...</v></c>
    for (auto j = 0; j < cell_val.length(); ++j) {
      std::string c_val = cell_val[j];
      std::string c_nam = cell_val_names[j];
      
      // Rcpp::Rcout << c_val << std::endl;
      
      // <f> ... </f>
      if(c_nam.compare("f") == 0) {
        cell.append_child(c_nam.c_str()).append_child(pugi::node_pcdata).set_value(c_val.c_str());
      }
      
      // <is><t> ... </t></is>
      if(c_nam.compare("is") == 0) {
        cell.append_child(c_nam.c_str()).append_child("t").append_child(pugi::node_pcdata).set_value(c_val.c_str());
      }
        
      // <v> ... </v>
      if(c_nam.compare("v") == 0) {
        cell.append_child(c_nam.c_str()).append_child(pugi::node_pcdata).set_value(c_val.c_str());
      }
      
    }
    
  }
  
  std::ostringstream oss;
  doc.print(oss, " ", pugi::format_raw);
  // doc.print(oss);
  
  return oss.str();
}

// [[Rcpp::export]]
Rcpp::CharacterVector set_sst(Rcpp::CharacterVector sharedStrings) {

  Rcpp::CharacterVector sst(sharedStrings.length());
  
    for (auto i = 0; i < sharedStrings.length(); ++i) {
      pugi::xml_document si;
      std::string sharedString = Rcpp::as<std::string>(sharedStrings[i]);
      
      si.append_child("si").append_child("t").append_child(pugi::node_pcdata).set_value(sharedString.c_str());
      
      std::ostringstream oss;
      si.print(oss, " ", pugi::format_raw);
      
      sst[i] = oss.str();
    }
    
    return sst;
}
