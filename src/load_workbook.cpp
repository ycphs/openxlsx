#include "openxlsx.h"


// [[Rcpp::export]]
SEXP getNodes(std::string xml, std::string tagIn){
  
  // This function loops over all characters in xml, looking for tag
  // tag should look liked <tag>
  // tagEnd is then generated to be <tag/>
  
  
  if(xml.length() == 0)
    return wrap(NA_STRING);
  
  xml = " " + xml;
  std::vector<std::string> r;
  size_t pos = 0;
  size_t endPos = 0;
  std::string tag = tagIn;
  std::string tagEnd = tagIn.insert(1,"/");
  
  size_t k = tag.length();
  size_t l = tagEnd.length();
  
  while(1){
    
    pos = xml.find(tag, pos+1);
    endPos = xml.find(tagEnd, pos+k);
    
    if((pos == std::string::npos) | (endPos == std::string::npos))
      break;
    
    r.push_back(xml.substr(pos, endPos-pos+l).c_str());
    
  }  
  
  CharacterVector out = wrap(r);  
  return markUTF8(out);
}


// [[Rcpp::export]]
SEXP getOpenClosedNode(std::string xml, std::string open_tag, std::string close_tag){
  
  if(xml.length() == 0)
    return wrap(NA_STRING);
  
  xml = " " + xml;
  size_t pos = 0;
  size_t endPos = 0;
  
  size_t k = open_tag.length();
  size_t l = close_tag.length();
  
  std::vector<std::string> r;
  
  while(1){
    
    pos = xml.find(open_tag, pos+1);
    endPos = xml.find(close_tag, pos+k);
    
    if((pos == std::string::npos) | (endPos == std::string::npos))
      break;
    
    r.push_back(xml.substr(pos, endPos-pos+l).c_str());
    
  }  
  
  CharacterVector out = wrap(r);  
  return markUTF8(out);
}


// [[Rcpp::export]]
SEXP getAttr(CharacterVector x, std::string tag){
  
  size_t n = x.size();
  size_t k = tag.length();
  
  if(n == 0)
    return wrap(-1);
  
  std::string xml;
  CharacterVector r(n);
  size_t pos = 0;
  size_t endPos = 0;
  std::string rtagEnd = "\"";
  
  for(size_t i = 0; i < n; i++){ 
    
    // find opening tag     
    xml = x[i];
    pos = xml.find(tag, 0);
    
    if(pos == std::string::npos){
      r[i] = NA_STRING;
    }else{  
      endPos = xml.find(rtagEnd, pos+k);
      r[i] = xml.substr(pos+k, endPos-pos-k).c_str();
    }
  }
  
  return markUTF8(r);   // no need to wrap as r is already a CharacterVector
  
}


// [[Rcpp::export]]
std::vector<std::string> getChildlessNode_ss(std::string xml, std::string tag){
  
  size_t k = tag.length();
  std::vector<std::string> r;
  size_t pos = 0;
  size_t endPos = 0;
  std::string tagEnd = "/>";
  
  while(1){
    
    pos = xml.find(tag, pos+1);    
    if(pos == std::string::npos)
      break;
    
    endPos = xml.find(tagEnd, pos+k);
    
    r.push_back(xml.substr(pos, endPos-pos+2).c_str());
    
  }
  
  return r ;  
  
}


// [[Rcpp::export]]
CharacterVector getChildlessNode(std::string xml, std::string tag) {
  
  size_t k = tag.length();
  if(xml.length() == 0)
    return wrap(NA_STRING);
  
  size_t begPos = 0, endPos = 0;
  
  std::vector<std::string> r;
  std::string res = "";
  
  // check "<tag "
  std::string begTag = "<" + tag + " ";
  std::string endTag = ">";
  
  // initial check, which kind of tags to expect
  begPos = xml.find(begTag, begPos);
  
  // if begTag was found
  if(begPos != std::string::npos) {
    
    endPos = xml.find(endTag, begPos);
    res = xml.substr(begPos, (endPos - begPos) + endTag.length());
    
    // check if last 2 characters are "/>"
    // <foo/> or <foo></foo>
    if (res.substr( res.length() - 2 ).compare("/>") != 0) {
      // check </tag>
      endTag = "</" + tag + ">";
    }
    
    // try with <foo ... />
    while( 1 ) {
      
      begPos = xml.find(begTag, begPos);
      endPos = xml.find(endTag, begPos);
      
      if(begPos == std::string::npos) 
        break;
      
      // read from initial "<" to final ">"
      res = xml.substr(begPos, (endPos - begPos) + endTag.length());
      
      begPos = endPos + endTag.length();
      r.push_back(res);
    }
  }
  
  
  CharacterVector out = wrap(r);  
  return markUTF8(out);
  
}


// [[Rcpp::export]]
CharacterVector get_extLst_Major(std::string xml){
  
  // find page margin or pagesetup then take the extLst after that
  
  if(xml.length() == 0)
    return wrap(NA_STRING);
  
  std::vector<std::string> r;
  std::string tagEnd = "</extLst>";
  size_t endPos = 0;
  std::string node;
  
  
  size_t pos = xml.find("<pageSetup ", 0);   
  if(pos == std::string::npos)
    pos = xml.find("<pageMargins ", 0);   
  
  if(pos == std::string::npos)
    pos = xml.find("</conditionalFormatting>", 0);   
  
  if(pos == std::string::npos)
    return wrap(NA_STRING);
  
  while(1){
    
    pos = xml.find("<extLst>", pos + 1);  
    if(pos == std::string::npos)
      break;
    
    endPos = xml.find(tagEnd, pos + 8);
    
    node = xml.substr(pos + 8, endPos - pos - 8);
    //pos = xml.find("conditionalFormattings", pos + 1);  
    //if(pos == std::string::npos)
    //  break;
    
    r.push_back(node.c_str());
    
  }
  
  CharacterVector out = wrap(r);  
  return markUTF8(out);
  
}


// [[Rcpp::export]]
int cell_ref_to_col( std::string x ){
  
  // This function converts the Excel column letter to an integer
  char A = 'A';
  int a_value = (int)A - 1;
  int sum = 0;
  
  // remove digits from string
  x.erase(std::remove_if(x.begin()+1, x.end(), ::isdigit),x.end());
  int k = x.length();
  
  for (int j = 0; j < k; j++){
    sum *= 26;
    sum += (x[j] - a_value);
  }
  
  return sum;
  
}


// [[Rcpp::export]]
CharacterVector int_2_cell_ref(IntegerVector cols){
  
  std::vector<std::string> LETTERS = get_letters();
  
  int n = cols.size();  
  CharacterVector res(n);
  std::fill(res.begin(), res.end(), NA_STRING);
  
  int x;
  int modulo;

  
  for(int i = 0; i < n; i++){
    
    if(!IntegerVector::is_na(cols[i])){

      string columnName;
      x = cols[i];
      while(x > 0){  
        modulo = (x - 1) % 26;
        columnName = LETTERS[modulo] + columnName;
        x = (x - modulo) / 26;
      }
      res[i] = columnName;
    }
    
  }
  
  return res ;
  
}
