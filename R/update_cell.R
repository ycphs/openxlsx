
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
#'    xlsxFile <- system.file("extdata", "update_test.xlsx", package = "openxlsx")
#'    wb <- loadWorkbook(xlsxFile)
#'    wb <- update_cell(x = c(1:3), wb = wb, sheet = "Sheet1", cell = "D4:D6")
#'    wb <- update_cell(x = 7, wb = wb, sheet = "Sheet1", cell = "H1")
#' @export
update_cell <- function(x, wb, sheet, cell) {
  
  dimensions <- unlist(strsplit(cell, ":"))
  rows <- gsub("[[:upper:]]","", dimensions)
  cols <- gsub("[[:digit:]]","", dimensions)
  
  if (length(dimensions) == 2) {
    # cols
    cols <- col2int(cols)
    cols <- seq(cols[1], cols[2])
    cols <- int2col(cols)
    
    rows <- as.character(seq(min(rows), max(rows)))
  }
  

  if(is.character(sheet)) {
    sheet_id <- which(sheet == wb$sheet_names)
  } else {
    sheet_id <- sheet
  }
  
  # if(identical(sheet_id, integer(0)))
  #   stop("sheet not in workbook")

  # pull sheet to modify from workbook; modify it; push back and write
  vval  <- wb$worksheets[[sheet_id]]$sheet_data$vval
  vtyp  <- wb$worksheets[[sheet_id]]$sheet_data$vtyp
  fval  <- wb$worksheets[[sheet_id]]$sheet_data$fval
  ftyp  <- wb$worksheets[[sheet_id]]$sheet_data$ftyp
  
  rtyp  <- wb$worksheets[[sheet_id]]$sheet_data$rtyp  # to identify dates?
  styp  <- wb$worksheets[[sheet_id]]$sheet_data$styp  # to identify dates?
  ttyp  <- wb$worksheets[[sheet_id]]$sheet_data$ttyp  # to identify strings and numbers
  isval <- wb$worksheets[[sheet_id]]$sheet_data$isval # inlinestr
  
  
  # workbooks contain only entries for values currently present.
  # if A1 is filled, B1 is not filled and C1 is filled the sheet will only
  # contain fields A1 and C1.
  cells_in_wb <- as.character(unlist(rtyp))
  rows_in_wb <- names(rtyp)
  
  
  if(!any(dimensions %in% cells_in_wb))
    stop("cell not in workbook")
  
  
  cn <- function(x) {
    x <- character(0)
    attr(x, "names") <- character(0)
    x
  }
  
  if (any(rows %in% rows_in_wb) )
    message("found cell(s) to update")
  
  if (all(rows %in% rows_in_wb)) {
    message("cell(s) to update already in workbook. updating ...")
    
    i <- 0
    for (row in rows) {
      for (col in cols) {
        i <- i+1
        vval[[row]][[col]] <- as.character(x[i])
        
        # for now create a str
        if (is.character(x)) {
          ttyp[[row]][[col]] <- "str"
        } else {
          ttyp[[row]][[col]] <- ""
        }
        
        vtyp[[row]][[col]] <- list(cn())
        
      }
    }
    
  }
  
  
  # push everything back to workbook
  wb$worksheets[[sheet_id]]$sheet_data$vval  <- vval
  wb$worksheets[[sheet_id]]$sheet_data$fval  <- fval
  wb$worksheets[[sheet_id]]$sheet_data$isval <- isval
  wb$worksheets[[sheet_id]]$sheet_data$rtyp  <- rtyp
  wb$worksheets[[sheet_id]]$sheet_data$styp  <- styp
  wb$worksheets[[sheet_id]]$sheet_data$ttyp  <- ttyp
  
  wb
}