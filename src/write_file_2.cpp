

#include "openxlsx.h"


// [[Rcpp::export]]
SEXP write_worksheet_xml_2( std::string prior
                          , std::string post
                          , Reference sheet_data
                          , Nullable<CharacterVector> row_heights_ = R_NilValue
                          , Nullable<CharacterVector> outline_levels_ = R_NilValue
                          , std::string R_fileName = "output"){
  

  // open file and write header XML
  const char * s = R_fileName.c_str();
  std::ofstream xmlFile;
  xmlFile.open (s);
  xmlFile << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
  xmlFile << prior;
  
  IntegerVector cell_row = sheet_data.field("rows");
  
  // If no data write childless node and return
  if(cell_row.size() == 0){

    xmlFile << "<sheetData/>";
    xmlFile << post;
    xmlFile.close();
    return Rcpp::wrap(0);
  }
  

  // sheet_data will be in order, just need to check for row_heights
  CharacterVector cell_col = int_2_cell_ref(sheet_data.field("cols"));
  CharacterVector cell_types = map_cell_types_to_char(sheet_data.field("t"));
  CharacterVector cell_value = sheet_data.field("v");
  CharacterVector cell_fn = sheet_data.field("f");
  
  CharacterVector style_id = sheet_data.field("style_id");
  CharacterVector unique_rows(sort_unique(cell_row));

  CharacterVector row_heights;
  CharacterVector row_heights_rows;
  size_t n_row_heights = 0;

  CharacterVector outline_levels;
  CharacterVector outline_levels_rows;
  CharacterVector outline_levels_hidden;
  size_t n_outline_levels = 0;

  if (row_heights_.isNotNull()) {
    row_heights = row_heights_;
    row_heights_rows = row_heights.attr("names");
    n_row_heights = row_heights.size();
  }

  if (outline_levels_.isNotNull()) {
    outline_levels = outline_levels_;
    outline_levels_rows = outline_levels.attr("names");
    outline_levels_hidden = outline_levels.attr("hidden");
    n_outline_levels = outline_levels.size();
  }

  size_t n = cell_row.size();
  size_t k = unique_rows.size();
  std::string xml;
  std::string cell_xml;
  
  size_t j = 0;
  size_t h = 0;
  size_t l = 0;
  String current_row = unique_rows[0];
  bool row_has_data = true;
  
  xmlFile << "<sheetData>";
  
  for(size_t i = 0; i < k; i++){
    
    cell_xml = "";
    row_has_data = false;
    
    while(current_row == unique_rows[i]){
      
      row_has_data = true;
      j += 1;
      
      if(CharacterVector::is_na(cell_col[j-1])){ //If r IS NA we have no row data we only have a rowHeight

        row_has_data = false;
        if(j == n)
          break;
        
        current_row = cell_row[j];
        break;
      }

      //cell XML strings
      cell_xml += "<c r=\"" + cell_col[j-1] + itos(cell_row[j-1]);
      
      if(!CharacterVector::is_na(style_id[j-1]))
        cell_xml += "\" s=\"" + style_id[j-1];
      
      
      //If we have a t value we must have a v value
      if(!CharacterVector::is_na(cell_types[j-1])){
        
        //If we have a c value we might have an f value
        if(CharacterVector::is_na(cell_fn[j-1])){ // no function: v or is
          if(Rcpp::as<std::string>(cell_types[j-1]).compare("inlineStr") == 0){
            cell_xml += "\" t=\"" + cell_types[j-1] + "\">" + "<is><t>" + cell_value[j-1] + "</t></is></c>";
          }else{
            cell_xml += "\" t=\"" + cell_types[j-1] + "\"><v>" + cell_value[j-1] + "</v></c>";
          }
        }else{
          if(CharacterVector::is_na(cell_value[j-1])){ // If v is NA
            cell_xml += "\" t=\"" + cell_types[j-1] + "\">" + cell_fn[j-1] + "</c>";
          }else{
            cell_xml += "\" t=\"" + cell_types[j-1] + "\">" + cell_fn[j-1] + "<v>" + cell_value[j-1] + "</v></c>";
          }
        }
        
        
      }else if(!CharacterVector::is_na(cell_fn[j-1])){
        cell_xml += "\">" + cell_fn[j-1] + "</c>";
      }else{
        cell_xml += "\"/>";
      }
      
      if(j == n)
        break;
      
      current_row = cell_row[j];
      
    }

    // <row..> element has the following structure, where most components are optional:
    //    <row r="1" ht="..." customHeight="1" outline_level="1" hidden="1">.....</row>

    // Handle one attribute after the other:
    xmlFile << "<row r=\"" + unique_rows[i] + "\"";
    // If there are custom row heights
    if ((h < n_row_heights) && (!Rf_isNull(row_heights_)) && (unique_rows[i] == row_heights_rows[h])) {
      xmlFile << " ht=\"" + row_heights[h] + "\" customHeight=\"1\"";
      h++;
    }
    // If there are grouped rows
    if ((l < n_outline_levels) && (!Rf_isNull(outline_levels_)) && (unique_rows[i] == outline_levels_rows[l])) {
      xmlFile << " outlineLevel=\"" + outline_levels[l] + "\"";
      // Ignore empty / null / NA values for hidden (accessing NULL would crash R / Rcpp!)
      if ((outline_levels_hidden[l] != R_NilValue) && (outline_levels_hidden[l] != NA_STRING) && (outline_levels_hidden[l] != "")) {
        xmlFile << " hidden=\"" + outline_levels_hidden[l] + "\"";
      }
      l++;
    }
    // If the row has contents
    if (row_has_data) {
      xmlFile << ">" + cell_xml + "</row>";
    } else {
      xmlFile << "/>";
    }
    
  }
  
  // All remaining grouped rows without content (ie. rows beyond the entries in unique_rows) need to be appended
  if ((!Rf_isNull(outline_levels_))) {
    while (l < n_outline_levels) {
      xmlFile << "<row r=\"" + outline_levels_rows[l] + "\" outlineLevel=\"" + outline_levels[l] + "\"";
      if ((outline_levels_hidden[l] != R_NilValue) && (outline_levels_hidden[l] != NA_STRING) && (outline_levels_hidden[l] != "")) {
        xmlFile << " hidden=\"" + outline_levels_hidden[l] + "\"";
      }
      xmlFile << "/>";
      l++;
    }
  }
    
  
  
  // write closing tag and XML post data
  xmlFile << "</sheetData>";
  xmlFile << post;
  
  //close file
  xmlFile.close();
  
  return wrap(0);
  
}
