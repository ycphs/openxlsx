#' dummy function to write data
#' @param wb workbook
#' @param sheet sheet
#' @param data data to export
#' @param colNames include colnames?
#' @param rowNames include rownames?
#' @param startRow row to place it
#' @param startCol col to place it
#' 
#' @examples 
#' # create a workbook and add some sheets
#' wb <- createWorkbook()
#' 
#' addWorksheet(wb, "sheet1")
#' writeData2(wb, "sheet1", mtcars, colNames = TRUE, rowNames = TRUE)
#' 
#' addWorksheet(wb, "sheet2")
#' writeData2(wb, "sheet2", cars, colNames = FALSE)
#' 
#' addWorksheet(wb, "sheet3")
#' writeData2(wb, "sheet3", letters)
#' 
#' addWorksheet(wb, "sheet4")
#' writeData2(wb, "sheet4", as.data.frame(Titanic), startRow = 2, startCol = 2)
#' 
#' saveWorkbook(wb, file = "/tmp/test.xlsx", overwrite = TRUE)
#'
#' @export
writeData2 <-function(wb, sheet, data,
                      colNames = TRUE, rowNames = FALSE,
                      startRow = 1, startCol = 1) {
  
  
  is_data_frame <- FALSE
  data_class <- sapply(data, class)
  
  # convert factor to character
  if (any(data_class == "factor")) {
    fcts <- which(data_class == "factor")
    data[fcts] <- lapply(data[fcts],as.character)
  }
  
  if (class(data) == "data.frame" | class(data) == "matrix") {
    is_data_frame <- TRUE
    
    # add colnames
    if (colNames == TRUE)
      data <- rbind(colnames(data), data)
    
    # add rownames
    if (rowNames == TRUE) {
      nam <- names(data)
      data <- cbind(rownames(data), data)
      names(data) <- c("", nam)
      data_class <- c("character", data_class)
    }
  }
  
  
  sheetno <- wb$validateSheet(sheet)
  # message("sheet no: ", sheetno)
  
  
  
  sheet_data <- list()
  
  # create a data frame
  if (!is_data_frame)
    data <- as.data.frame(t(data))
  
  data_nrow <- NROW(data)
  data_ncol <- NCOL(data)
  
  endRow <- (startRow -1) + data_nrow
  endCol <- (startCol -1) + data_ncol
  
  
  dims <- paste0(int2col(startCol), startRow,
                 ":",
                 int2col(endCol), endRow)
  
  wb$worksheets[[sheetno]]$dimension <- paste0("<dimension ref=\"", dims, "\"/>")
  
  # rtyp character vector per row 
  # list(c("A1, ..., "k1"), ...,  c("An", ..., "kn"))
  rtyp <- openxlsx:::dims_to_dataframe(dims, fill = TRUE)
  
  rows_attr <- cols_attr <- cc <- vector("list", data_nrow)
  
  cols_attr <- lapply(seq_len(data_nrow),
                      function(x) list(collapsed="false",
                                       customWidth="true",
                                       hidden="false",
                                       outlineLevel="0",
                                       max="121",
                                       min="1",
                                       style="0",
                                       width="9.14"))
  
  wb$worksheets[[sheetno]]$cols_attr <- openxlsx:::list_to_attr(cols_attr, "col")
  
  
  
  rows_attr <- lapply(startRow:endRow,
                      function(x) list("r" = as.character(x),
                                       "spans" = paste0("1:", data_ncol),
                                       "x14ac:dyDescent"="0.25"))
  names(rows_attr) <- rownames(rtyp)
  
  wb$worksheets[[sheetno]]$sheet_data$row_attr <- rows_attr
  
  
  
  
  numcell <- function(x,y){
    list(val = list(v = as.character(x)),
         typ = list(r = y),
         attr = list(empty = "empty"))
  }
  
  chrcell <- function(x,y){
    list(val = list(v = x),
         typ = list(r = y, t = "str"),
         attr = list(empty = "empty"))
  }
  
  cell <- function(x, y, data_class) {
    z <- NULL
    if (data_class == "numeric")
      z <- numcell(x,y)
    if (data_class %in% c("character", "factor"))
      z <- chrcell(x,y)
    
    z
  }
  
  
  for (i in seq_len(nrow(data))) {
  
      col <- vector("list", ncol(data))
      for (j in seq_along(data)) {
        dc <- ifelse(colNames && i == 1, "character", data_class[j])
        col[[j]] <- cell(data[i, j], rtyp[i, j], dc)
      }
      names(col) <- colnames(rtyp)
      
      cc[[i]] <- col
  }
  names(cc) <- rownames(rtyp)
  
  wb$worksheets[[sheetno]]$sheet_data$cc <- cc
  
  # update a few styles informations
  wb$styles$numFmts <- character(0)
  wb$styles$cellXfs <- "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
  wb$styles$dxfs <- character(0)
  
  wb 
}
