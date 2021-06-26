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
SEXP getXMLXPtr2val(XPtrXML doc, std::string level1, std::string child) {
  
  // returns a single vector, not a list of vectors!
  std::vector<std::string> x;
  
  for (pugi::xml_node worksheet = doc->child(level1.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(level1.c_str()))
  {
    std::vector<std::string> y;
    
    pugi::xml_node col = worksheet.child(child.c_str());
    x.push_back(col.child_value() );
    
  }
  
  return  Rcpp::wrap(x);
}

// [[Rcpp::export]]
SEXP getXMLXPtr3val(XPtrXML doc, std::string level1, std::string level2, std::string child) {
  
  // returns a single vector, not a list of vectors!
  std::vector<std::string> x;
  
  for (pugi::xml_node worksheet = doc->child(level1.c_str()).child(level2.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(level2.c_str()))
  {
    std::vector<std::string> y;
    
    pugi::xml_node col = worksheet.child(child.c_str());
    
    x.push_back(col.child_value() );
  }
  
  return  Rcpp::wrap(x);
}

// [[Rcpp::export]]
SEXP getXMLXPtr4val(XPtrXML doc, std::string level1, std::string level2, std::string level3, std::string child) {
  
  // returns a list of vectors!
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
      y.push_back(col.child_value() );
    }
    
    x.push_back(y);
  }
  
  return  Rcpp::wrap(x);
}

// [[Rcpp::export]]
SEXP getXMLXPtr5val(XPtrXML doc, std::string level1, std::string level2, std::string level3, std::string level4, std::string child) {
  
  auto worksheet = doc->child(level1.c_str()).child(level2.c_str());
  size_t n = std::distance(worksheet.begin(), worksheet.end());
  auto itr_rows = 0;
  Rcpp::List x(n);
  
  for (pugi::xml_node worksheet = doc->child(level1.c_str()).child(level2.c_str()).child(level3.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(level3.c_str()))
  {
    size_t k = std::distance(worksheet.begin(), worksheet.end());
    auto itr_cols = 0;
    Rcpp::List y(k);
    
    std::vector<std::string> nam;
    
    for (pugi::xml_node col = worksheet.child(level4.c_str());
         col;
         col = col.next_sibling(level4.c_str()))
    {
      Rcpp::CharacterVector z;
      
      // get r attr e.g. "A1"
      std::string colrow = col.attribute("r").value();
      // remove numeric from string
      colrow.erase(std::remove_if(colrow.begin(),
                                  colrow.end(),
                                  &isdigit),
                                  colrow.end());
      nam.push_back(colrow);
      
      for (pugi::xml_node val = col.child(child.c_str());
           val;
           val = val.next_sibling(child.c_str()))
      {
        std::string val_s = "";
        // is node contains additional t node.
        // TODO: check if multiple t nodes are possible, for now return only one.
        if (val.child("t")) {
          pugi::xml_node tval = val.child("t");
          val_s = tval.child_value();
        } else {
          val_s = val.child_value();
        }
        
        z.push_back( val_s );
      }
      
      y[itr_cols]= z;
      ++itr_cols;
    }
    
    y.attr("names") = nam;
    
    x[itr_rows] = y;
    ++itr_rows;
  }
  
  return  Rcpp::wrap(x);
}

// [[Rcpp::export]]
SEXP getXMLXPtr1attr(XPtrXML doc, std::string child) {
  
  
  pugi::xml_node worksheet = doc->child(child.c_str());
  size_t n = std::distance(worksheet.begin(), worksheet.end());
  
  Rcpp::List z(n);
  
  auto itr = 0;
  for (worksheet = doc->child(child.c_str());
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

// [[Rcpp::export]]
SEXP getXMLXPtr2attr(XPtrXML doc, std::string level1, std::string child) {
  
  
  pugi::xml_node worksheet = doc->child(level1.c_str()).child(child.c_str());
  size_t n = std::distance(worksheet.begin(), worksheet.end());
  
  Rcpp::List z(n);
  
  auto itr = 0;
  for (worksheet = doc->child(level1.c_str()).child(child.c_str());
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
  
  auto rows = doc->child(level1.c_str()).child(level2.c_str());
  size_t n = std::distance(rows.begin(), rows.end());
  auto itr_rows = 0;
  Rcpp::List z(n);
  
  for (pugi::xml_node worksheet = doc->child(level1.c_str()).child(level2.c_str()).child(level3.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(level3.c_str()))
  {
    size_t k = std::distance(worksheet.begin(), worksheet.end());
    auto itr_cols = 0;
    Rcpp::List y(k);
    
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
        if (attr.value() != NULL) {
          nam.push_back(attr.name());
          res.push_back(attr.value());
        } else {
          res.push_back("");
        }
      }
      
      // assign names
      res.attr("names") = nam;
      
      y[itr_cols] = res;
      ++itr_cols;
    }
    
    z[itr_rows] = y;
    ++itr_rows;
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
  
  auto worksheet = doc->child(level1.c_str()).child(level2.c_str());
  size_t n = std::distance(worksheet.begin(), worksheet.end());
  auto itr_cols = 0;
  Rcpp::List z(n);
  
  for (pugi::xml_node worksheet = doc->child(level1.c_str()).child(level2.c_str()).child(level3.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(level3.c_str()))
  {
    
    size_t k = std::distance(worksheet.begin(), worksheet.end());
    auto itr_rows = 0;
    Rcpp::List y(k);
    
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
        
        for (pugi::xml_attribute attr = col.first_attribute();
             attr;
             attr = attr.next_attribute())
        {
          if (attr.value() != NULL) {
            nam.push_back(attr.name());
            res.push_back(attr.value());
          } else {
            res.push_back("");
          }
        }
        
        // assign names
        res.attr("names") = nam;
        
        x.push_back(res);
      }
      
      y[itr_rows] = x;
      ++itr_rows;
    }
    
    z[itr_cols] = y;
    ++itr_cols;
  }
  
  return  Rcpp::wrap(z);
}

// [[Rcpp::export]]
Rcpp::CharacterVector getXMLXPtr1attr_one(XPtrXML doc, std::string child, std::string attrname) {
  
  
  std::vector<std::string> z;
  
  for (pugi::xml_node worksheet = doc->child(child.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(child.c_str()))
  { 
    pugi::xml_attribute attr = worksheet.attribute(attrname.c_str());
    
    if (attr.value() != NULL) {
      z.push_back(attr.value());
    } else {
      z.push_back("");
    }
  }
  
  return  Rcpp::wrap(z);
}

// [[Rcpp::export]]
Rcpp::CharacterVector getXMLXPtr2attr_one(XPtrXML doc, std::string level1, std::string child, std::string attrname) {
  
  
  std::vector<std::string> z;
  
  for (pugi::xml_node worksheet = doc->child(level1.c_str()).child(child.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(child.c_str()))
  { 
    pugi::xml_attribute attr = worksheet.attribute(attrname.c_str());
    
    if (attr.value() != NULL) {
      z.push_back(attr.value());
    } else {
      z.push_back("");
    }
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
    
    if (attr.value() != NULL) {
      z.push_back(attr.value());
    } else {
      z.push_back("");
    }
  }
  
  return  Rcpp::wrap(z);
}


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
// [[Rcpp::export]]
SEXP getXMLXPtr4attr_one(XPtrXML doc, std::string level1, std::string level2, std::string level3, std::string child, std::string attrname) {
  
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
      
      if (attr.value() != NULL) {
        y.push_back(attr.value());
      } else {
        y.push_back("");
      }
    }
    
    z.push_back(y);
  }
  
  return  Rcpp::wrap(z);
}

// [[Rcpp::export]]
std::string printXPtr(XPtrXML doc, bool raw) {
  
  std::ostringstream oss;
  if (raw) {
    doc->print(oss, " ", pugi::format_raw);
  } else {
    doc->print(oss);
  }
  
  return  oss.str();
}
