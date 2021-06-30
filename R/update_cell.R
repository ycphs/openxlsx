
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
#'    wb <- update_cell(x = c("x", "y", "z"), wb = wb, sheet = "Sheet1", cell = "B3:D3")
#'    wb <- update_cell(x = 7, wb = wb, sheet = "Sheet1", cell = "D4")
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

  # 1) pull sheet to modify from workbook; 2) modify it; 3) push it back
  cc  <- wb$worksheets[[sheet_id]]$sheet_data$cc
  
  
  # workbooks contain only entries for values currently present.
  # if A1 is filled, B1 is not filled and C1 is filled the sheet will only
  # contain fields A1 and C1.
  cells_in_wb <- as.character(unlist(lapply(cc, function(x) sapply(x, function(y)y[["typ"]]$r))))
  rows_in_wb <- names(cc)
  
  
  if(!any(dimensions %in% cells_in_wb))
    stop("cell not in workbook")
  
  
  if (any(rows %in% rows_in_wb) )
    message("found cell(s) to update")
  
  if (all(rows %in% rows_in_wb)) {
    message("cell(s) to update already in workbook. updating ...")
    
    i <- 0
    for (row in rows) {
      for (col in cols) {
        i <- i+1
        
        cc[[row]][[col]][["val"]]$f <- NULL
        cc[[row]][[col]][["typ"]]$s <- NULL
        cc[[row]][[col]][["typ"]]$t <- NULL
        cc[[row]][[col]][["val"]]$v <- NULL
        cc[[row]][[col]][["val"]]$is <- NULL
        cc[[row]][[col]][["attr"]]$t <- NULL
        cc[[row]][[col]][["attr"]]$ref <- NULL
        cc[[row]][[col]][["attr"]]$si <- NULL
        
        # for now create a str
        if (is.character(x)) {
          cc[[row]][[col]][["typ"]]$t <- "inlineStr"
          cc[[row]][[col]][["val"]]$is <- as.character(x[i])
        } else {
          cc[[row]][[col]][["val"]]$v <- as.character(x[i])
        }
        
        cc[[row]][[col]][["attr"]]$empty <- "empty"
        
      }
    }
    
  }
  
  
  # push everything back to workbook
  wb$worksheets[[sheet_id]]$sheet_data$cc  <- cc
  
  wb
}
