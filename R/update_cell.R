
#' replace data in a single cell
#' 
#' Minimal invasive update of a single cell for imported workbooks.
#' 
#' @param x value you want to insert
#' @param wb the workbook you want to update
#' @param sheet the sheet you want to update
#' @param cell the cell you want to update in Excel conotation e.g. "A1"
#' 
#' @examples 
#' \dontrun{
#'    wb <- update_cell(x = "Updated 2021-04-02", wb = wb, sheet = "Lookups", cell = "G1")
#'    wb <- update_cell(x = 7, wb = wb, sheet = "Lookups", cell = "H1")
#' }
#' @export
update_cell <- function(x, wb, sheet, cell) {

  if(is.character(sheet)) {
    sheet_id <- which(sheet == wb$sheet_names)
  } else {
    sheet_id <- sheet
  }
  
  cells <- wb$worksheets[[sheet_id]]$sheet_data$rtyp
  to_update <- lapply(cells, function(x) cell == x)
  
  
  cn <- function(x) {
    x <- character(0)
    attr(x, "names") <- character(0)
    x
  }
  
  any_to_update <- any(unlist(to_update))
  
  if (any_to_update) {
    message("found cell to update")
   
    row_to_update <- which(unlist(lapply(to_update, any)))
    col_to_update <- which(to_update[[row_to_update]])
   
    # # get the character
    # col <- gsub("[^a-zA-Z]", "", cell)
    
    wb$worksheets[[sheet_id]]$sheet_data$vval[[row_to_update]][[col_to_update]] <- as.character(x)
    
    # for now create a str
    if (is.character(x)) {
      wb$worksheets[[sheet_id]]$sheet_data$ttyp[[row_to_update]][[col_to_update]] <- "str"
    } else {
      wb$worksheets[[sheet_id]]$sheet_data$ttyp[[row_to_update]][[col_to_update]] <- ""
    }
    
    # if vtyp character or numeric
    # must be a list with a named numeric, if not, its broken
    wb$worksheets[[sheet_id]]$sheet_data$vtyp[[row_to_update]][[col_to_update]] <- list(cn())
  }
  
  wb
}