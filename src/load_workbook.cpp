
#include "openxlsx.h"



// [[Rcpp::export]]
SEXP loadworksheets(Reference wb, List styleObjects, std::vector<std::string> xmlFiles, LogicalVector is_chart_sheet){
  
  List worksheets = wb.field("worksheets");
  int n_sheets = is_chart_sheet.size();
  CharacterVector sheetNames = wb.field("sheet_names");
  
  // variable set up
  std::string tagEnd = "\"";
  std::string cell;
  List colWidths(n_sheets);
  List rowHeights(n_sheets);
  List wbstyleObjects(0);
  List outlineLevels(n_sheets);
  List colOutlineLevels(n_sheets);

  // loop over each worksheet file
  for(int i = 0; i < n_sheets; i++){
    
    if(is_chart_sheet[i]){
      
      colWidths[i] = List(0);
      rowHeights[i] = List(0);
      outlineLevels[i] = List(0);
      colOutlineLevels[i] = List(0);
      
    }else{
      
      colWidths[i] = List(0);
      rowHeights[i] = List(0);
      outlineLevels[i] = List(0);
      colOutlineLevels[i] = List(0);
      Reference this_worksheet(worksheets[i]);
      Reference sheet_data(this_worksheet.field("sheet_data"));
      
      //read in file
      std::string xmlFile = xmlFiles[i];
      
      std::string buf;
      std::string xml = read_file_newline(xmlFile);
      // ifstream file;
      // file.open(xmlFile.c_str());
      // while (file >> buf)
      // xml += buf + ' ';
      
      size_t pos = xml.find("<sheetData>");  // find <sheetData>
      size_t endPos = 0;
      size_t tmp_pos = 0;
      
      bool has_data = true;
      if(pos == string::npos){
        has_data = false;
        pos = xml.find("<sheetData/>");
        if(pos == string::npos){
          pos = xml.find("<sheetData />");
        }
      }
      
      /* --- Everything before pos --- */
      std::string xml_pre = xml.substr(0, pos);
      
      // sheetPR
      CharacterVector sheetPr = getNodes(xml_pre, "<sheetPr>");
      
      if(sheetPr.size() == 0){
        sheetPr = getNodes(xml_pre, "<sheetPr");
        
        for(int j = 0; j < sheetPr.size(); j++){
          std::string sp = as<std::string>(sheetPr[j]);
          char ch = *sp.rbegin();  
          if(ch != '>')
            sp += ">";
          sheetPr[j] = sp;
        }
      }
      
      if(sheetPr.size() == 0)
        sheetPr = getChildlessNode(xml_pre, "sheetPr");
      
      if(sheetPr.size() > 0)
        this_worksheet.field("sheetPr") = sheetPr;
      
      
      
      // Freeze Panes
      CharacterVector node_xml = getChildlessNode(xml_pre, "pane");
      if(node_xml.size() > 0)
        this_worksheet.field("freezePane") = node_xml;
      
      // SheetViews
      node_xml = getNodes(xml_pre, "<sheetViews>");
      if(node_xml.size() > 0)
        this_worksheet.field("sheetViews") = node_xml;
      
      
      //colwidths
      std::vector<std::string> cols = getChildlessNode_ss(xml_pre, "<col ");
      if(cols.size() > 0){
        
        NumericVector widths;
        IntegerVector columns_with_widths;
        IntegerVector columns_with_groups;
        CharacterVector column_hidden;
        CharacterVector col_outline;
        CharacterVector col_hidden;
        
        for (size_t ci = 0; ci < cols.size(); ci++) {
          
          double tmp_width = 0;
          std::string tmp_hidden;
          int min_c = 0;
          int max_c = 0;
          std::string tmp_coloutline;

          buf = cols[ci];
          // If either custom widths or groupings, get column index
          if ((buf.find("customWidth", 0) != string::npos) | (buf.find("outlineLevel", 0) != string::npos)) {
            
            tmp_pos = buf.find("min=\"", 0);
            endPos = buf.find(tagEnd, tmp_pos + 5);
            min_c = atoi(buf.substr(tmp_pos + 5, endPos - tmp_pos - 5).c_str());
            
            tmp_pos = buf.find("max=\"", 0);
            endPos = buf.find(tagEnd, tmp_pos + 5);
            max_c = atoi(buf.substr(tmp_pos + 5, endPos - tmp_pos - 5).c_str());

            tmp_pos = buf.find("hidden=\"", 0);

            if (tmp_pos != string::npos) {
              endPos = buf.find(tagEnd, tmp_pos + 8);
              tmp_hidden = buf.substr(tmp_pos + 8, endPos - tmp_pos - 8);
            } else {
              tmp_hidden = "0";
            }

            // If column has both a custom width and is part of a group
            if ((buf.find("customWidth", 0) != string::npos) & (buf.find("outlineLevel", 0) != string::npos)) {
              tmp_pos = buf.find("width=\"", 0);
              endPos = buf.find(tagEnd, tmp_pos + 7);
              tmp_width = atof(buf.substr(tmp_pos + 7, endPos - tmp_pos - 7).c_str()) - 0.71;
              
              tmp_pos = buf.find("outlineLevel=\"", 0);
              endPos = buf.find(tagEnd, tmp_pos + 14);
              tmp_coloutline = buf.substr(tmp_pos + 14, endPos - tmp_pos - 14);  

              while (min_c <= max_c) {
                widths.push_back(tmp_width);
                columns_with_widths.push_back(min_c);
                columns_with_groups.push_back(min_c);
                column_hidden.push_back(tmp_hidden);
                col_hidden.push_back(tmp_hidden);
                col_outline.push_back(tmp_coloutline);
                min_c++;
              }
              
            } else if (buf.find("customWidth", 0) != string::npos) {      // Column only has a custom width

              tmp_pos = buf.find("width=\"", 0);
              endPos = buf.find(tagEnd, tmp_pos + 7);
              tmp_width = atof(buf.substr(tmp_pos + 7, endPos - tmp_pos - 7).c_str()) - 0.71;

              while (min_c <= max_c) {
                widths.push_back(tmp_width);
                columns_with_widths.push_back(min_c);
                column_hidden.push_back(tmp_hidden);
                min_c++;
              }
            } else {      // Column is only part of a group

              tmp_pos = buf.find("outlineLevel=\"", 0);
              endPos = buf.find(tagEnd, tmp_pos + 14);
              tmp_coloutline = buf.substr(tmp_pos + 14, endPos - tmp_pos - 14);  

              while (min_c <= max_c) {
                columns_with_groups.push_back(min_c);
                col_hidden.push_back(tmp_hidden);
                col_outline.push_back(tmp_coloutline);
                min_c++;
              }
            }
          }
        }
        
        if (widths.size() > 0) {
          CharacterVector tmp_widths(widths);
          tmp_widths.attr("names") = columns_with_widths;
          tmp_widths.attr("hidden") = column_hidden;
          colWidths[i] = tmp_widths;
        }

        if (col_outline.size() > 0) {
          CharacterVector columns_outline(col_outline);
          columns_outline.attr("names") = columns_with_groups;
          columns_outline.attr("hidden") = col_hidden;
          colOutlineLevels[i] = columns_outline;
        }
      }
      
      /* --- Everything after sheetData --- */
      size_t pos_post = 0;
      if(has_data){
        pos_post = xml.find("</sheetData>");   
      }else{
        pos_post = pos; 
      }
      
      std::string xml_post = xml.substr(pos_post);
      
      node_xml = getChildlessNode(xml_post, "sheetProtection");
      if(node_xml.size() > 0) {
        this_worksheet.field("sheetProtection") = node_xml;
      }
      
      
      node_xml = getChildlessNode(xml_post, "autoFilter");
      if(node_xml.size() > 0)
        this_worksheet.field("autoFilter") = node_xml;
      
      
      node_xml = getChildlessNode(xml_post, "hyperlink");
      if(node_xml.size() > 0)
        this_worksheet.field("hyperlinks") = node_xml;
      
      
      node_xml = getChildlessNode(xml_post, "pageMargins");
      if(node_xml.size() > 0)
        this_worksheet.field("pageMargins") = node_xml;
      
      
      node_xml = getChildlessNode(xml_post, "pageSetup");
      if(node_xml.size() > 0){
        for(int j = 0; j < node_xml.size(); j++){
          
          std::string pageSetup_tmp = as<std::string>(node_xml[j]);
          size_t ps_pos = pageSetup_tmp.find("r:id=\"rId", 0);
          if(ps_pos != std::string::npos){
            
            std::string pageSetup_tmp2 = pageSetup_tmp.substr(0, ps_pos + 9) + "2";
            ps_pos = pageSetup_tmp.find("\"", ps_pos + 9);
            pageSetup_tmp  = pageSetup_tmp2 + pageSetup_tmp.substr(ps_pos);
            
          }
          
          node_xml[j] = pageSetup_tmp;
          
        }
        this_worksheet.field("pageSetup") = node_xml;
      }
      
      node_xml = getChildlessNode(xml_post, "mergeCell");
      if(node_xml.size() > 0)
        this_worksheet.field("mergeCells") = node_xml;
      
      
      node_xml = getNodes(xml_post, "<oleObjects>");
      if(node_xml.size() > 0)
        this_worksheet.field("oleObjects") = node_xml;
      
      
      // headerfooter
      CharacterVector xml_hf = getNodes(xml_post, "<headerFooter");
      if(xml_hf.size() > 0){
        
        List hf = List(0);
        
        node_xml = getNodes(xml_post, "<oddHeader>");
        if(node_xml.size() > 0)
          hf["oddHeader"] = node_xml;
        
        node_xml = getNodes(xml_post, "<oddFooter>");
        if(node_xml.size() > 0)
          hf["oddFooter"] = node_xml;
        
        node_xml = getNodes(xml_post, "<evenHeader>");
        if(node_xml.size() > 0)
          hf["evenHeader"] = node_xml;
        
        node_xml = getNodes(xml_post, "<evenFooter>");
        if(node_xml.size() > 0)
          hf["evenFooter"] = node_xml;
        
        node_xml = getNodes(xml_post, "<firstHeader>");
        if(node_xml.size() > 0)
          hf["firstHeader"] = node_xml;
        
        node_xml = getNodes(xml_post, "<firstFooter>");
        if(node_xml.size() > 0)
          hf["firstFooter"] = node_xml;
        
        this_worksheet.field("headerFooter") = hf;
        
      }
      
      
      node_xml = getChildlessNode(xml_post, "drawing");
      if(node_xml.size() == 0)
        node_xml = getChildlessNode(xml_post, "legacyDrawing");
      
      if(node_xml.size() > 0){
        for(int j = 0; j < node_xml.size(); j++){
          
          std::string drawingId_tmp = as<std::string>(node_xml[j]);
          size_t ps_pos = drawingId_tmp.find("r:id=\"rId", 0);
          
          std::string drawingId_tmp2 = drawingId_tmp.substr(0, ps_pos + 9) + "1";
          ps_pos = drawingId_tmp.find("\"", ps_pos + 9);
          
          drawingId_tmp  = drawingId_tmp2 + drawingId_tmp.substr(ps_pos);
          node_xml[j] = drawingId_tmp;
          
        }
      }
      
      
      //  conditionalFormatting
      CharacterVector conForm = getNodes(xml_post, "<conditionalFormatting");
      if(conForm.size() > 0){
        
        // get sqref attribute
        size_t tmp_pos = 0;
        int end_pos = 0;
        std::string sqref;
        CharacterVector cf;
        CharacterVector cf_names;
        
        
        for(int ci = 0; ci < conForm.size(); ci++){
          
          buf = conForm[ci];
          
          tmp_pos = buf.find("sqref=\"", 0);
          end_pos = buf.find("\"", tmp_pos + 7);
          
          sqref = buf.substr(tmp_pos + 7, end_pos - tmp_pos - 7);
          buf = buf.substr(0, buf.find("</conditionalFormatting"));
          buf = buf.substr(buf.find("<cfRule"));
          
          int ocs = 0;
          string::size_type start = 0;
          while((start = buf.find("<cfRule", start)) != string::npos){
            ++ocs;
            start += 7;
          }
          
          if(ocs == 1){
            cf_names.push_back(sqref);
            cf.push_back(buf);
          }else if(ocs > 1){
            
            tmp_pos = buf.find("<cfRule", 0);
            while(tmp_pos != std::string::npos){
              
              end_pos = buf.find("<cfRule", tmp_pos + 7);
              cf.push_back(buf.substr(tmp_pos, end_pos - tmp_pos));
              cf_names.push_back(sqref);
              tmp_pos = end_pos;
              
            }
          }
          
        } // end of loop through conditional formats
        
        cf.attr("names") = cf_names;
        this_worksheet.field("conditionalFormatting") = cf;
        
      } // end of if(conForm.size() > 0)
      
      
      //data validation
      node_xml = getOpenClosedNode(xml_post, "<dataValidation ", "</dataValidation>");
      if(node_xml.size() > 0)
        this_worksheet.field("dataValidations") = node_xml;
      
      // extLst
      node_xml = get_extLst_Major(xml_post);
      if(node_xml.size() > 0)
        this_worksheet.field("extLst") = node_xml;
      
      
      // clean pre and post xml
      xml_post.clear();
      xml_pre.clear();
      
      
      /* --------------------------- sheet Data --------------------------- */
      
      if(has_data){
        
        xml = xml.substr(pos + 11, pos_post - pos - 11);     // get from "<sheetData>" to the end
        
        // count cells with children
        int ocs = 0;
        string::size_type start = 0;
        while((start = xml.find("<c ", start)) != string::npos){
          ++ocs;
          start += 4;
        }
        
        CharacterVector r(ocs);
        CharacterVector r_nms(ocs);
        CharacterVector v(ocs);
        CharacterVector s(ocs);
        
        
        std::fill(v.begin(), v.end(), NA_STRING);
        std::fill(s.begin(), s.end(), NA_STRING);
        
        // rebuild
        CharacterVector t(ocs);
        std::fill(t.begin(), t.end(), "n");
        
        IntegerVector rows_cell_ref(ocs);
        IntegerVector cols_cell_ref(ocs);
        
        CharacterVector f(ocs);
        std::fill(f.begin(), f.end(), NA_STRING);
        // rebuild end
        
        int j = 0;
        size_t nextPos = 3;
        pos = xml.find("<c ", 0);
        bool has_v = false;
        bool has_f = false;
        std::string func;
        
        size_t pos_t = pos;
        size_t pos_f = pos;
        
        // PULL OUT CELL AND ATTRIBUTES
        while(j < ocs){
          
          if(pos != std::string::npos){
            
            has_v = false;
            has_f = false;
            
            nextPos = xml.find("<c ", pos + 9);
            cell = xml.substr(pos, nextPos - pos);
            
            // Pull out ref
            pos = cell.find("r=", 0);  // find r="
            endPos = cell.find(tagEnd, pos + 3);  // find next "
            r[j] = cell.substr(pos + 3, endPos - pos - 3).c_str();
            
            buf = cell.substr(pos + 3, endPos - pos - 3);      
            cols_cell_ref[j] = cell_ref_to_col(buf);
            
            buf.erase(std::remove_if(buf.begin(), buf.end(), ::isalpha), buf.end());
            r_nms[j] = buf;
            
            rows_cell_ref[j] = atoi(buf.c_str());
            
            
            // Pull out style
            pos = cell.find(" s=", 0);  // find s="
            if(pos != std::string::npos){
              endPos = cell.find(tagEnd, pos + 4);  // find next "
              s[j] = cell.substr(pos + 4, endPos - pos - 4);
            }
            
            // find <v> tag and </v> end tag
            endPos = cell.find("</v>", 0);
            if(endPos != std::string::npos){
              pos = cell.find("<v", 0);
              pos = cell.find(">", pos);
              v[j] = cell.substr(pos + 1, endPos - pos - 1);
              has_v = true;
            }
            
            // find <is><t> tag and </t></is> end tag
            endPos = cell.find("</t></is>", 0);
            if(endPos != std::string::npos){
              pos = cell.find("<is><t", 0);
              pos = cell.find(">", pos);
              v[j] = cell.substr(pos + 4, endPos - pos - 4); // skip <t> and </t
              has_v = true;
            }
            
            // Pull out type
            pos_t = cell.find(" t=", 0);
            pos_f = cell.find("<f", 0);
            
            // have both
            if((pos_f != std::string::npos) & (pos_t != std::string::npos)){ // have f
              
              
              // will always have f
              endPos = cell.find("</f>", pos_f + 3);
              if(endPos == std::string::npos){
                endPos = cell.find("/>", pos_f + 3);
                f[j] = cell.substr(pos_f, endPos - pos_f + 2);
              }else{
                f[j] = cell.substr(pos_f, endPos - pos_f + 4);
              }
              has_f = true;
              
              // do we really have t
              if(pos_t < pos_f){
                endPos = cell.find(tagEnd, pos_t + 4);  // find next "
                t[j] = cell.substr(pos_t + 4, endPos - pos_t - 4);
              }
              
              
            }else if(pos_t != std::string::npos){ // only have t
              
              endPos = cell.find(tagEnd, pos_t + 4);  // find next "
              t[j] = cell.substr(pos_t + 4, endPos - pos_t - 4);
              
              
            }else if(pos_f != std::string::npos){ // only have f
              
              endPos = cell.find("</f>", pos_f + 3);
              if(endPos == std::string::npos){
                endPos = cell.find("/>", pos_f + 3);
                f[j] = cell.substr(pos_f, endPos - pos_f + 2);
              }else{
                f[j] = cell.substr(pos_f, endPos - pos_f + 4);
              }
              has_f = true;
              
            }
            
            
            if(has_f & (!has_v) & (t[j] != "n")){
              
              v[j] = NA_STRING;
              
            }else if(has_f & !has_v){
              
              t[j] = NA_STRING;
              v[j] = NA_STRING;
              
            }else if(has_f | has_v){
 
            }else{ //only have s and r
              t[j] = NA_STRING;
              v[j] = NA_STRING;
            }
            
            
            
            j++; // INCREMENT OVER OCCURENCES
            pos = nextPos;
            pos_t = nextPos;
            pos_f = nextPos;
            
            
          }  // end of while loop over occurences
        }  // END OF CELL AND ATTRIBUTION GATHERING
        
        // get names of cells
        
        if(ocs > 0){
          
          // may be a problem when we have a formula, no value and we now write t="n" in it's place
          sheet_data.field("rows") = rows_cell_ref;
          sheet_data.field("cols") = cols_cell_ref;
          
          sheet_data.field("t") = map_cell_types_to_integer(t);
          sheet_data.field("v") = v;
          sheet_data.field("f") = f;
          
          sheet_data.field("data_count") = 1;
          sheet_data.field("n_elements") = ocs;
          
        }
        
        // count number of rows
        int row_ocs = 0;
        start = 0;
        while((start = xml.find("<row ", start)) != string::npos){
          ++row_ocs;
          start += 4;
        }
        
        CharacterVector rowNumbers(row_ocs);
        CharacterVector heights(row_ocs);
        CharacterVector outlines(row_ocs);
        CharacterVector outline_hidden(row_ocs);
        
        
        // PULL OUT CELL AND ATTRIBUTES
        j = 0;
        pos = xml.find("<row ", 0);
        std::string htTag = " ht=\"";
        std::string attrEnd = "\"";
        std::string posVal;
        
        while(j < row_ocs){
          
          if(pos != std::string::npos){
            
            nextPos = xml.find("<row ", pos + 9);
            cell = xml.substr(pos, nextPos - pos);
            
            
            // Pull out ref
            pos = cell.find("r=", 0);  // find r="
            endPos = cell.find(tagEnd, pos + 3);  // find next "
            rowNumbers[j] = cell.substr(pos + 3, endPos - pos - 3);
            
            
            // find custom height  
            pos = cell.find(htTag, pos);
            if(pos == std::string::npos){
              heights[j] = NA_STRING;
            }else{  
              endPos = cell.find(attrEnd, pos + 5);
              heights[j] = cell.substr(pos + 5, endPos - pos - 5);
            };

            // Is row grouped?
            pos = cell.find("outlineLevel=\"", 0);
            if(pos == std::string::npos){
              outlines[j] = NA_STRING;
            } else {
              endPos = cell.find(attrEnd, pos + 14);
              posVal = cell.substr(pos + 14, endPos - pos - 14);
              // outlines[j] = posVal;
              // Ignore if value is 0
              if (posVal != "0") {
                outlines[j] = posVal;
              } else {
                outlines[j] = NA_STRING;
              }
            };
            
            // Is grouped row hidden?
            pos = cell.find("hidden=\"", 0);
            if(pos == std::string::npos){
              // outline_hidden[j] = "0";   // NA_STRING;
              outline_hidden[j] = NA_STRING;
            } else {
              endPos = cell.find(attrEnd, pos + 8);
              posVal = cell.substr(pos + 8, endPos - pos - 8);
              outline_hidden[j] = posVal;
              // Ignore if there's no grouping for that row
              if (!CharacterVector::is_na(outlines[j])) {
                outline_hidden[j] = posVal;
              } else {
                outline_hidden[j] = NA_STRING;
              }
            };
            
            j++; // INCREMENT OVER OCCURENCES
            pos = nextPos;
            
          }  // end of while loop over occurences
        }  // END OF CELL AND ATTRIBUTION GATHERING
        
        
        CharacterVector heightsRows(rowNumbers);
        CharacterVector outlineRows(rowNumbers);

        heightsRows = heightsRows[!is_na(heights)];
        if (heightsRows.size() > 0) {
        	heights = heights[!is_na(heights)];
        	heights.attr("names") = heightsRows;
        	rowHeights[i] = heights;
        }

        outlineRows = outlineRows[!is_na(outlines)];
        if (outlineRows.size() > 0) {
          // If a row is grouped, we need to preserve the visibility flag, even
          // if the cell is visible and the flag is NA. Otherwise, hidden
          // rows/columns would be shifted or saving would break (Issue
          // ycphs/openxlsx#138)
          outline_hidden = outline_hidden[!is_na(outlines)];
          outlines = outlines[!is_na(outlines)];
          outlines.attr("names") = outlineRows;
          outlines.attr("hidden") = outline_hidden;
          outlineLevels[i] = outlines;
        }


        
        // styleObjects
        std::string this_sheetname = as<std::string>(sheetNames[i]);
        
        if(any(!is_na(s))){
          
          CharacterVector s_refs = r[!is_na(s)];
          s = s[!is_na(s)];
          
          CharacterVector uStyleInds = sort_unique(s);
          int nsu = uStyleInds.size();
          CharacterVector uStyleInds_j(1);
          
          std::string ref_j;
          CharacterVector styleElementNames = CharacterVector::create("style", "sheet", "rows", "cols");
          
          for(int j = 0; j < nsu; j++){
            
            List styleElement(4);
            int styleInd = atoi(as<std::string>(uStyleInds[j]).c_str());
            
            if(styleInd != 0){
              
              uStyleInds_j[0] = uStyleInds[j];
              LogicalVector ind = !is_na(match(s, uStyleInds_j));
              CharacterVector s_refs_j = s_refs[ind];
              
              int n_j = s_refs_j.size();
              IntegerVector rows(n_j);
              IntegerVector cols = convert_from_excel_ref(s_refs_j);
              
              for(int k = 0; k < n_j; k++){
                ref_j = s_refs_j[k];
                ref_j.erase(std::remove_if(ref_j.begin(), ref_j.end(), ::isalpha), ref_j.end());
                rows[k] = atoi(ref_j.c_str());  
              }
              
              styleElement[0] = styleObjects[styleInd - 1];
              styleElement[1] = this_sheetname;
              styleElement[2] = rows;
              styleElement[3] = cols;
              
              styleElement.attr("names") = styleElementNames;
              
              wbstyleObjects.push_back(styleElement);
              
            }
            
          }
          
          
        } // end if(any(!is_na(s)))
        
      } // end of if(has_data)
      
    } // end if is_chart_sheet[i] else
    
  } // end of loop over sheets
  
  // assign back to workbook
  wb.field("worksheets") = worksheets;
  wb.field("rowHeights") = rowHeights;
  wb.field("colWidths") = colWidths;
  wb.field("styleObjects") = wbstyleObjects;
  wb.field("outlineLevels") = outlineLevels;
  wb.field("colOutlineLevels") = colOutlineLevels;

  
  return wrap(wb);
  
}








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

  // size_t k = tag.length(); variable not used
  if(xml.length() == 0)
    return wrap(NA_STRING);

  size_t begPos = 0, endPos = 0;

  std::vector<std::string> r;
  std::string res = "";

  // check "<tag "
  const std::string begTag = "<" + tag;
  const std::string endTag = ">";
  std::string temp_endTag = "";

  // initial check, which kind of tags to expect
  begPos = xml.find(begTag, begPos);

  // if begTag was found
  // try with <foo ... />
  while( begPos != std::string::npos ) {

    endPos = xml.find(endTag, begPos);
    if(begPos == std::string::npos || endPos == std::string::npos) break;
    res = xml.substr(begPos, (endPos - begPos) + endTag.length());
    if (res.length() == 0) break;

    auto itr = 0;
    // check if we have either <foo ...>, <foo/>, or <foo>. We have to avoid
    // <foos...>
    while (
        res.substr(begTag.length(),1).compare(" ") != 0 &&   // <foo ...>
          res.substr(begTag.length(),1).compare("/") != 0 && // <foo/>
          res.substr(begTag.length(),1).compare(">") != 0    // <foo>
    ) {
      
      if (itr == 0) begPos = begPos + begTag.length();
      if(begPos == std::string::npos || endPos == std::string::npos) break;

      Rcpp::checkUserInterrupt();

      begPos = xml.find(begTag, begPos);
      endPos = xml.find(endTag, begPos);

      if(begPos == std::string::npos || endPos == std::string::npos) break;
      res = xml.substr(begPos, (endPos - begPos) + endTag.length());

      ++itr;
    }

    // if we have <foo> we need to find the matching closing tag </foo>
    bool closingtag = false;
    if (res.substr( res.length() - 2 ).compare("/>") != 0) {
      // this node has </tag>
      temp_endTag = "</" + tag + ">";
      closingtag = true;
    } else {
      temp_endTag = endTag;
    }

    // if we have a closing tag, we need to reposition the endPos. Previously
    // it was at the end of <foo>. Now we search for </foo>
    if (closingtag) {
      endPos = xml.find(temp_endTag, begPos);
      if(begPos == std::string::npos || endPos == std::string::npos) break;
      
      // read from initial "<" to final ">"
      res = xml.substr(begPos, (endPos - begPos) + temp_endTag.length());
      if (res.length() == 0) break;
    }

    if(begPos == std::string::npos || endPos == std::string::npos) break;

    begPos = endPos + temp_endTag.length();
    begPos = xml.find(begTag, begPos);
    r.push_back(res);

    if(begPos == std::string::npos || endPos == std::string::npos) break;
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
