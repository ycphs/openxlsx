
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
#'    
#'    wb <- loadWorkbook(xlsxFile)
#'    wb <- update_cell(x = 7, wb = wb, sheet = "Sheet1", cell = "A4")
#'    
#'    wb_to_df(wb)
#'    
#' @export
update_cell <- function(x, wb, sheet, cell, data_class, colNames = FALSE) {
  
  dimensions <- unlist(strsplit(cell, ":"))
  rows <- gsub("[[:upper:]]","", dimensions)
  cols <- gsub("[[:digit:]]","", dimensions)
  
  if (length(dimensions) == 2) {
    # cols
    cols <- col2int(cols)
    cols <- seq(cols[1], cols[2])
    cols <- int2col(cols)
    
    rows <- as.character(seq(rows[1], rows[2]))
  }
  

  if(is.character(sheet)) {
    sheet_id <- which(sheet == wb$sheet_names)
  } else {
    sheet_id <- sheet
  }
  
  if (missing(data_class))
    data_class <- sapply(x, class)
  
  # if(identical(sheet_id, integer(0)))
  #   stop("sheet not in workbook")

  # 1) pull sheet to modify from workbook; 2) modify it; 3) push it back
  cc  <- wb$worksheets[[sheet_id]]$sheet_data$cc
  row_attr <- wb2$worksheets[[sheet_id]]$sheet_data$row_attr
  
  # workbooks contain only entries for values currently present.
  # if A1 is filled, B1 is not filled and C1 is filled the sheet will only
  # contain fields A1 and C1.
  cells_in_wb <- as.character(unlist(lapply(cc, function(x) sapply(x, function(y)y[["typ"]]$r))))
  rows_in_wb <- names(cc)
  

  # check if there are rows not available
  if (!all(rows %in% rows_in_wb)) {
    # message("row(s) not in workbook")
  
    # add row to name vector, extend the entire thing
    total_rows <- as.character(sort(unique(as.numeric(c(rows, rows_in_wb)))))
    
    # new list
    cc_new <- vector("list", length(total_rows))
    names(cc_new) <- total_rows
    
    # new row_attr
    row_attr_new <- vector("list", length(rows_in_wb))
    names(row_attr_new) <- rows_in_wb
    
    row_attr_new[rows_in_wb] <- row_attr[rows_in_wb]
    
    for (trow in total_rows) {
      row_attr_new[[trow]] <- list(r = trow)
    }
    
    wb$worksheets[[sheet_id]]$sheet_data$row_attr <- row_attr_new
    
    # assign what we already have
    cc_new[rows_in_wb] <- cc
    
    # provide output
    cc <- cc_new
    rows_in_wb <- total_rows
    
  }
  
  if (!any(cols %in% cells_in_wb)) {
    # all rows are availabe in the list
    for (row in rows){
      
      # collect all wanted cols and order for excel
      total_cols <- unique(c(names(cc[[row]]), cols))
      total_cols <- int2col(sort(col2int(total_cols)))
      
      # create candidate
      cc_row_new <- vector("list", length(total_cols))
      names(cc_row_new) <- total_cols
      
      # extract row (easier or maybe only way to change order?)
      cc_row <- cc[[row]]
      cc_row_new[names(cc[[row]])] <- cc_row[names(cc[[row]])]
      
      # assign to cc
      cc[[row]] <- cc_row_new
      
      # insert empty for newly created cells
      for (col in cols) {
        cc[[row]][[col]] <- list(val = list(empty = "empty"),
                                 typ = list(r = paste0(col, row)),
                                 attr = list(empty = "empty"))
      }
    }
  }
  
  # update dimensions
  cells_in_wb <- as.character(unlist(lapply(cc, function(x) sapply(x, function(y)y[["typ"]]$r))))
  
  all_rows <- as.integer(unique(gsub("[[:upper:]]","", cells_in_wb)))
  all_cols <- unique(gsub("[[:digit:]]","", cells_in_wb))
  all_cols <- ifelse(nchar(all_cols) == 1, paste0(" ", all_cols), all_cols)
  
  min_cell <- trimws(paste0(min(all_cols), min(all_rows)))
  max_cell <- trimws(paste0(max(all_cols), max(all_rows)))
  
  # print(min_cell)
  # print(max_cell)
  # print(rows)
  # print(cols)
  
  # i know, i know, i'm lazy
  wb$worksheets[[sheet_id]]$dimension <- paste0("<dimension ref=\"", min_cell, ":", max_cell, "\"/>")
  
  # if (any(rows %in% rows_in_wb) )
    # message("found cell(s) to update")
  
  if (all(rows %in% rows_in_wb)) {
    # message("cell(s) to update already in workbook. updating ...")
    
    i <- 0; n <- 0
    for (row in rows) {
      
      n <- n+1
      m <- 0
      
      for (col in cols) {
        i <- i+1
        m <- m+1
        
        # check if is data frame or matrix
        value <- ifelse(is.null(dim(x)), x[i], x[n, m])
        
        
        cc[[row]][[col]][["val"]]$empty <- NULL
        
        cc[[row]][[col]][["val"]]$f <- NULL
        cc[[row]][[col]][["typ"]]$s <- NULL
        cc[[row]][[col]][["typ"]]$t <- NULL
        cc[[row]][[col]][["val"]]$v <- NULL
        cc[[row]][[col]][["val"]]$is <- NULL
        cc[[row]][[col]][["attr"]]$t <- NULL
        cc[[row]][[col]][["attr"]]$ref <- NULL
        cc[[row]][[col]][["attr"]]$si <- NULL
        
        # for now create a str
        if (data_class[m] %in% c("character", "factor") | (colNames == TRUE & n == 1)) {
          cc[[row]][[col]][["typ"]]$t <- "inlineStr"
          cc[[row]][[col]][["val"]]$is <- as.character(value)
        } else {
          cc[[row]][[col]][["val"]]$v <- as.character(value)
        }
        
        cc[[row]][[col]][["attr"]]$empty <- "empty"
        
      }
    }
    
  }
  
  
  # push everything back to workbook
  wb$worksheets[[sheet_id]]$sheet_data$cc  <- cc
  
  wb
}
