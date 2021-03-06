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
  
  rows_attr <- cols_attr <- vtyp <- vector("list", data_nrow)
  
  cn <- function(x) {
    x <- character(0)
    attr(x, "names") <- character(0)
    x
  }
  
  
  cols_attr <- lapply(seq_len(data_nrow),
                      function(x) c(collapsed="false",
                                    customWidth="true",
                                    hidden="false",
                                    outlineLevel="0",
                                    max="121",
                                    min="1",
                                    style="0",
                                    width="9.14"))
  
  wb$worksheets[[sheetno]]$cols_attr <- cols_attr
  
  
  
  rows_attr <- lapply(startRow:endRow,
                      function(x) c("r" = as.character(x),
                                    "spans" = paste0("1:", data_ncol),
                                    "x14ac:dyDescent"="0.25"))
  
  wb$worksheets[[sheetno]]$rows_attr <- rows_attr
  
  # fval list per row
  fval <- lapply(seq_len(data_nrow),
                 function(x) 
                   lapply(seq_len(data_ncol),
                          function(x)
                            character(0)))
  
  wb$worksheets[[sheetno]]$sheet_data$fval <- fval
  
  # ftyp list per row
  ftyp <-  lapply(seq_len(data_nrow),
                  function(x) 
                    lapply(seq_len(data_ncol),
                           function(x) 
                             list()))
  
  wb$worksheets[[sheetno]]$sheet_data$ftyp <- ftyp
  
  # rtyp character vector per row 
  # list(c("A1, ..., "k1"), ...,  c("An", ..., "kn"))
  rtyp <- lapply(startRow:endRow, function(x) 
    paste0(openxlsx:::convert_to_excel_ref(startCol:endCol, LETTERS), x))
  
  wb$worksheets[[sheetno]]$sheet_data$rtyp <- rtyp
  
  # styp character vector per row (style typ, choose 1)
  styp <- lapply(seq_len(data_nrow), function(x)
    rep("", data_ncol))
  
  
  wb$worksheets[[sheetno]]$sheet_data$styp <- styp
  
  
  xlsx_types <- function(x) {
    res <- NULL
    for (cls in x) {
      if (cls == "numeric" | cls == "integer")
        res <- c(res, "")
      if (cls == "character" | cls == "factor")
        res <- c(res, "str")
    }
    res
  }
  
  # ttyp character vector per row ("s", "str" or "" )
  ttyp <- lapply(seq_len(data_nrow), function(x)
    xlsx_types(data_class))
  
  strtyp <- "str"
  
  # for data frames set first row with names as str
  if (is_data_frame & colNames)
    ttyp[[1]] <- rep(strtyp, length(ttyp[[1]]))
  
  wb$worksheets[[sheetno]]$sheet_data$ttyp <- ttyp
  
  
  vval <- lapply(seq_len(data_nrow),
                 function(x) lapply(seq_len(data_ncol),
                                    function(i) as.character(data[x, i])))
  
  wb$worksheets[[sheetno]]$sheet_data$vval <- vval
  
  
  vtyp <- lapply(seq_len(data_nrow),
                 function(x) 
                   lapply(seq_len(data_ncol),
                          function(x) 
                            lapply(1,
                                   function(x)
                                     cn())))
  
  wb$worksheets[[sheetno]]$sheet_data$vtyp <- vtyp
  
  wb$worksheets[[sheetno]]$dimension <- paste0("<dimension ref=\"",
                                               openxlsx:::convert_to_excel_ref(startCol, LETTERS),
                                               startRow,
                                               ":",
                                               openxlsx:::convert_to_excel_ref(endCol, LETTERS),
                                               endRow, "\"/>")
  
  # update a few styles informations
  wb$styles$numFmts <- character(0)
  wb$styles$cellXfs <- "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
  wb$styles$dxfs <- character(0)
  
  wb 
}
