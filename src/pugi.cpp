#include "openxlsx_types.h"

// [[Rcpp::export]]
SEXP readXMLPtr(std::string path) {
  
  
  xmldoc *doc = new xmldoc;
  
  pugi::xml_parse_result result = doc->load_file(path.c_str(),
                                                 pugi::parse_default | pugi::parse_escapes);
  if (!result) {
    Rcpp::stop("xml import unsuccessfull");
  }
  
  XPtrXML ptr(doc, true);
  ptr.attr("class") = Rcpp::CharacterVector::create("pugi_xml");
  return ptr;
}


// [[Rcpp::export]]
SEXP getXMLXPtr1(XPtrXML doc, std::string child) {
  
  std::vector<std::string> res;
  
  for (pugi::xml_node worksheet = doc->child(child.c_str());
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
SEXP getXMLXPtr2(XPtrXML doc, std::string level1, std::string child) {
  
  std::vector<std::string> res;
  
  for (pugi::xml_node worksheet = doc->child(level1.c_str()).child(child.c_str());
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
SEXP getXMLXPtr3(XPtrXML doc, std::string level1, std::string level2, std::string child) {
  
  std::vector<std::string> res;
  
  for (pugi::xml_node worksheet = doc->child(level1.c_str()).child(level2.c_str()).child(child.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(child.c_str()))
  {
    std::ostringstream oss;
    worksheet.print(oss, " ", pugi::format_raw);
    res.push_back(oss.str());
  }
  
  return  Rcpp::wrap(res);
}

//nested list below level 3. eg:
//<level1>
//  <level2>
//    <level3>
//      <child />
//      x
//      <child />
//    </level3>
//    <level3>
//      <child />
//    </level3>
//  </level2>
//</level1>
// [[Rcpp::export]]
SEXP getXMLXPtr4(XPtrXML doc, std::string level1, std::string level2, std::string level3, std::string child) {
  
  std::vector<std::vector<std::string>> x;
  
  for (pugi::xml_node worksheet = doc->child(level1.c_str()).child(level2.c_str()).child(level3.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(level3.c_str()))
  {
    std::vector<std::string> y;
    
    for (pugi::xml_node col = worksheet.child(child.c_str());
         col;
         col = col.next_sibling(child.c_str()))
    {
      std::ostringstream oss;
      col.print(oss, " ", pugi::format_raw);
      
      y.push_back(oss.str());
    }
    
    x.push_back(y);
  }
  
  return  Rcpp::wrap(x);
}

// nested list below level 3. eg:
// <level1>
//  <level2>
//    <level3>
//      <level4 />
//        <child>
//        x
//        </child>
//        <child>
//        y
//        </child>
//      <level4 />
//    </level3>
//    <level3>
//      <level4 />
//    </level3>
//  </level2>
//</level1>
// [[Rcpp::export]]
SEXP getXMLXPtr5(XPtrXML doc, std::string level1, std::string level2, std::string level3, std::string level4, std::string child) {
  
  std::vector<std::vector<std::vector<std::string>>> x;
  
  for (pugi::xml_node worksheet = doc->child(level1.c_str()).child(level2.c_str()).child(level3.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(level3.c_str()))
  {
    std::vector<std::vector<std::string>> y;
    
    for (pugi::xml_node col = worksheet.child(level4.c_str());
         col;
         col = col.next_sibling(level4.c_str()))
    {
      std::vector<std::string> z;
      
      for (pugi::xml_node val = col.child(child.c_str());
           val;
           val = val.next_sibling(child.c_str()))
      {
        std::ostringstream oss;
        val.print(oss, " ", pugi::format_raw);
        z.push_back(oss.str());
      }
      
      y.push_back(z);
    }
    
    x.push_back(y);
  }
  
  return  Rcpp::wrap(x);
}


// [[Rcpp::export]]
SEXP getXMLXPtr3attr(XPtrXML doc, std::string level1, std::string level2, std::string child) {
  
  
  Rcpp::List z;
  
  for (pugi::xml_node worksheet = doc->child(level1.c_str()).child(level2.c_str()).child(child.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(child.c_str()))
  {
    
    Rcpp::CharacterVector res;
    std::vector<std::string> nam;
    
    for (pugi::xml_attribute attr = worksheet.first_attribute();
         attr;
         attr = attr.next_attribute())
    {
      nam.push_back(attr.name());
      res.push_back(attr.value());
    }
    
    // assign names
    res.attr("names") = nam;
    
    z.push_back(res);
  }
  
  return  Rcpp::wrap(z);
}


// nested list below level 3. eg:
// <worksheet>
//  <sheetData>
//    <row>
//      <c />
//      <c />
//    </row>
//    <row>
//      <c />
//    </row>
//  </sheetData>
//</worksheet>
// [[Rcpp::export]]
SEXP getXMLXPtr4attr(XPtrXML doc, std::string level1, std::string level2, std::string level3, std::string child) {
  
  Rcpp::List z;
  
  for (pugi::xml_node worksheet = doc->child(level1.c_str()).child(level2.c_str()).child(level3.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(level3.c_str()))
  {
    Rcpp::List y;
    
    for (pugi::xml_node row = worksheet.child(child.c_str());
         row;
         row = row.next_sibling(child.c_str()))
    {
      
      Rcpp::CharacterVector res;
      std::vector<std::string> nam;
      
      for (pugi::xml_attribute attr = row.first_attribute();
           attr;
           attr = attr.next_attribute())
      {
        nam.push_back(attr.name());
        res.push_back(attr.value());
      }
      
      // assign names
      res.attr("names") = nam;
      
      y.push_back(res);
    }
    z.push_back(y);
  }
  
  return  Rcpp::wrap(z);
}

// nested list below level 3. eg:
// <level1>
//  <level2>
//    <level3>
//      <level4 />
//        <child>
//        x
//        </child>
//        <child>
//        y
//        </child>
//      <level4 />
//    </level3>
//    <level3>
//      <level4 />
//    </level3>
//  </level2>
//</level1>
// [[Rcpp::export]]
SEXP getXMLXPtr5attr(XPtrXML doc, std::string level1, std::string level2, std::string level3, std::string level4, std::string child) {
  
  Rcpp::List z;
  
  for (pugi::xml_node worksheet = doc->child(level1.c_str()).child(level2.c_str()).child(level3.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(level3.c_str()))
  {
    Rcpp::List y;
    
    for (pugi::xml_node row = worksheet.child(level4.c_str());
         row;
         row = row.next_sibling(level4.c_str()))
    {
      Rcpp::List x;
      
      for (pugi::xml_node col = row.child(child.c_str());
           col;
           col = col.next_sibling(child.c_str()))
      {
        
        Rcpp::CharacterVector res;
        std::vector<std::string> nam;
        
        for (pugi::xml_attribute attr = row.first_attribute();
             attr;
             attr = attr.next_attribute())
        {
          nam.push_back(attr.name());
          res.push_back(attr.value());
        }
        
        // assign names
        res.attr("names") = nam;
        
        x.push_back(res);
      }
      
      y.push_back(x);
    }
    
    z.push_back(y);
  }
  
  return  Rcpp::wrap(z);
}

// [[Rcpp::export]]
Rcpp::CharacterVector getXMLXPtr3attr_one(XPtrXML doc, std::string level1, std::string level2, std::string child, std::string attrname) {
  
  
  std::vector<std::string> z;
  
  for (pugi::xml_node worksheet = doc->child(level1.c_str()).child(level2.c_str()).child(child.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(child.c_str()))
  { 
    pugi::xml_attribute attr = worksheet.attribute(attrname.c_str());
    z.push_back(attr.value());
  }
  
  return  Rcpp::wrap(z);
}

// [[Rcpp::export]]
SEXP getXMLXPtr4attr_one(XPtrXML doc, std::string level1, std::string level2, std::string level3, std::string child, std::string attrname) {
  // nested list below level 3. eg:
  // <level1>
  //  <level2>
  //    <level3>
  //      <child attrname=""/>
  //      <child />
  //    </level3>
  //    <level3>
  //      <child />
  //    </level3>
  //  </level2>
  //</level1
  
  std::vector<std::vector<std::string>> z;
  
  for (pugi::xml_node worksheet = doc->child(level1.c_str()).child(level2.c_str()).child(level3.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(level3.c_str()))
  {
    std::vector<std::string> y;
    
    for (pugi::xml_node row = worksheet.child(child.c_str());
         row;
         row = row.next_sibling(child.c_str()))
    {
      pugi::xml_attribute attr = row.attribute(attrname.c_str());
      y.push_back(attr.value());
    }
    
    z.push_back(y);
  }
  
  return  Rcpp::wrap(z);
}

// [[Rcpp::export]]
std::string printXPtr(XPtrXML doc) {
  
  std::ostringstream oss;
  doc->print(oss, " ", pugi::format_raw);
  
  return  oss.str();
}
