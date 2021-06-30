#include "openxlsx_types.h"

// [[Rcpp::export]]
std::string set_row(Rcpp::List row_attr, Rcpp::List cells) {
  
  pugi::xml_document doc;
  
  pugi::xml_node row = doc.append_child("row");
  Rcpp::CharacterVector attrnams = row_attr.names();
  
  for (auto i = 0; i < row_attr.length(); ++i) {
    row.append_attribute(attrnams[i]) = Rcpp::as<std::string>(row_attr[i]).c_str();
  }
  
  for (auto i = 0; i < cells.length(); ++i) {
    
    // create node <c>
    pugi::xml_node cell = row.append_child("c");
    
    Rcpp::List cll = cells[i];
    // Rf_PrintValue(cll);

    Rcpp::List cell_atr = cll["typ"];
    Rcpp::List cell_val = cll["val"];
    Rcpp::List attr_val = cll["attr"];
    
    // Rf_PrintValue(cell_atr);
    // Rf_PrintValue(cell_val);
    // Rf_PrintValue(attr_val);
    
    std::vector<std::string> cell_atr_names = cell_atr.names();
    std::vector<std::string> cell_val_names = cell_val.names();
    std::vector<std::string> attr_val_names = attr_val.names();
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
        
        pugi::xml_node f = cell.append_child(c_nam.c_str());
        
        for (auto k = 0; k < attr_val.length(); ++k) {
          std::string c_atr = attr_val_names[k];
          
          if (c_atr.compare("empty") != 0) {
            std::string c = attr_val[k];
            f.append_attribute(attr_val_names[k].c_str()) = c.c_str();
            f.set_value(c_val.c_str());
          }
        }
        
        f.append_child(pugi::node_pcdata).set_value(c_val.c_str());
        
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


// [[Rcpp::export]]
std::string list_to_attr(Rcpp::List attributes, std::string node) {
  
  pugi::xml_document doc;
  
  for (auto i = 0; i < attributes.length(); ++i) {
    pugi::xml_node nds = doc.append_child(node.c_str());
    
    Rcpp::List attrs = attributes[i];
    Rcpp::CharacterVector attrnams = attrs.names();
    for (auto j = 0; j < attrs.length(); ++j) {
      nds.append_attribute(attrnams[j]) = Rcpp::as<std::string>(attrs[j]).c_str();
    }
  }
  
  std::ostringstream oss;
  doc.print(oss, " ", pugi::format_raw);
  // doc.print(oss);
  
  return oss.str();
}


// [[Rcpp::export]]
std::string list_to_attr_full(Rcpp::List attributes, std::string node, std::string child) {
  
  pugi::xml_document doc;
  pugi::xml_node nds = doc.append_child(node.c_str());
  for (auto i = 0; i < attributes.length(); ++i) {
    nds.append_child(child.c_str());
    
    Rcpp::List attrs = attributes[i];
    Rcpp::CharacterVector attrnams = attrs.names();
    for (auto j = 0; j < attrs.length(); ++j) {
      nds.append_attribute(attrnams[j]) = Rcpp::as<std::string>(attrs[j]).c_str();
    }
  }
  
  std::ostringstream oss;
  doc.print(oss, " ", pugi::format_raw);
  // doc.print(oss);
  
  return oss.str();
}