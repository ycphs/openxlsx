#include <Rcpp.h>
#include <sstream>
#include "pugixml.hpp"

// [[Rcpp::export]]
SEXP getXMLattr(std::vector<std::string> strs, std::string child) {
  
  
  Rcpp::List z;
  
  
  for (const auto& str: strs) {
    
    
    
    Rcpp::CharacterVector res = {""};
    
    std::vector<std::string> nam = {""};
    
    
    if (str.compare("") != 0) {
      pugi::xml_document doc;
      
      pugi::xml_parse_result result = doc.load_string(str.c_str());
      if (!result) {
        Rcpp::stop("xml import unsuccessfull");
      }
      
      for (pugi::xml_attribute attr = doc.child(child.c_str()).first_attribute();
           attr;
           attr = attr.next_attribute())
      {
        nam.push_back(attr.name());
        res.push_back(attr.value());
      }
      
      // is it safe to search for next_child ?
      if (pugi::xml_node doc_ali = doc.child(child.c_str()).child("alignment")) {
        
        for (pugi::xml_attribute attr = doc_ali.first_attribute();
             attr;
             attr = attr.next_attribute())
        {
          nam.push_back(attr.name());
          res.push_back(attr.value());
        }
      }
      
      if (pugi::xml_node doc_pro = doc.child(child.c_str()).child("protection")) {
        
        for (pugi::xml_attribute attr = doc_pro.first_attribute();
             attr;
             attr = attr.next_attribute())
        {
          nam.push_back(attr.name());
          res.push_back(attr.value());
        }
      }
    }
    
    // assign names
    res.attr("names") = nam;
    
    z.push_back(res);
    
  }
  
  return  Rcpp::wrap(z);
}