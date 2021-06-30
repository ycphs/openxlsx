
#' create dataframe from dimensions
#' @param dims Character vector of expected dimension.
#' @param fill If TRUE, fills the dataframe with variables
#' @examples {
#'   dims_to_dataframe("A1:B2")
#' }
#' @export
dims_to_dataframe <- function(dims, fill = FALSE) {
  
  dimensions <- strsplit(dims, ":")
  rows <- as.numeric(gsub("[[:upper:]]","", dimensions[[1]]))
  cols <- gsub("[[:digit:]]","", dimensions[[1]])
  
  # cols
  col_min <- col2int(cols[1])
  col_max <- col2int(cols[2])
  
  cols <- seq(col_min, col_max)
  cols <- int2col(cols)
  
  # rows
  rows_min <- rows[1]
  rows_max <- rows[2]
  
  rows <- seq(rows_min, rows_max)
  
  data <- NA
  if (fill) {
    data <- expand.grid(cols, rows, stringsAsFactors = FALSE)
    data <- paste0(data$Var1, data$Var2)
  }
  
  # matrix as.data.frame
  mm <- matrix(data = data,
               nrow = length(rows),
               ncol = length(cols),
               dimnames = list(rows, cols), byrow = TRUE)
  
  z <- as.data.frame(mm)
  z
}

#' @param rtyp rtyp object
#' @export 
get_row_names <- function(rtyp) {
  sapply(rtyp, function(x) gsub("[[:upper:]]","", x[[1]]))
}

#' function to estimate the column type.
#' 0 = character, 1 = numeric, 2 = date.
#' @param tt dataframe produced by wb_to_df()
#' @export
guess_col_type <- function(tt) {
  
  # everythings character
  types <- vector("numeric", ncol(tt))
  names(types) <- names(tt)
  
  # but some values are numerics
  col_num <- sapply(tt, function(x) all(is.na(x)))
  types[names(col_num[col_num == TRUE])] <- 1
  
  # or even dates
  col_dte <- sapply(tt[!col_num], function(x) all(x[is.na(x) == FALSE] == "d"))
  types[names(col_dte[col_dte == TRUE])] <- 2
  
  types
}


#' Create Dataframe from Workbook
#' 
#' Simple function to create a dataframe from a workbook. Simple as in simply
#' written down and not optimized etc. The goal was to have something working.
#' 
#' @param xlsxFile An xlsx file, Workbook object or URL to xlsx file.
#' @param sheet Either sheet name or index. When missing the first sheet in the workbook is selected.
#' @param colnames If TRUE, the first row of data will be used as column names.
#' @param dims Character string of type "A1:B2" as optional dimentions to be imported.
#' @param detectDates If TRUE, attempt to recognise dates and perform conversion.
#' @param showFormula If TRUE, the underlying Excel formulas are shown.
#' @param convert If TRUE, a conversion to dates and numerics is attempted.
#' @param skipEmptyCols If TRUE, empty columns are skipped.
#' @param startRow first row to begin looking for data.
#' @param rows A numeric vector specifying which rows in the Excel file to read. If NULL, all rows are read.
#' @param cols A numeric vector specifying which columns in the Excel file to read. If NULL, all columns are read.
#' @param definedName Character string with a definedName. If no sheet is selected, the first appearance will be selected.
#' @param types A named numeric indicating, the type of the data. 0: character, 1: numeric, 2: date. Names must match the created
#' @param na.strings A character vector of strings which are to be interpreted as NA. Blank cells will be returned as NA.
#' @examples
#' 
#'   ###########################################################################
#'   # numerics, dates, missings, bool and string
#'   xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx")
#'   wb1 <- loadWorkbook(xlsxFile)
#' 
#'   # import workbook
#'   wb_to_df(wb1)
#'   
#'   # do not convert first row to colNames 
#'   wb_to_df(wb1, colNames = FALSE)
#'   
#'   # do not try to identify dates in the data
#'   wb_to_df(wb1, detectDates = FALSE)
#'   
#'   # return the underlying Excel formula instead of their values
#'   wb_to_df(wb1, showFormula = TRUE)
#'   
#'   # read dimension withot colNames
#'   wb_to_df(wb1, dims = "A2:C5", colNames = FALSE)
#'   
#'   # read selected cols
#'   wb_to_df(wb1, cols = c(1:2, 7))
#'   
#'   # read selected rows
#'   wb_to_df(wb1, rows = c(1, 4, 6))
#'   
#'   # convert characters to numerics and date (logical too?)
#'   wb_to_df(wb1, convert = FALSE)
#'   
#'   # erase empty Rows from dataset
#'   wb_to_df(wb1, sheet = 3, skipEmptyRows = TRUE)
#'   
#'   # erase rmpty Cols from dataset
#'   wb_to_df(wb1, skipEmptyCols = TRUE)
#'   
#'   # convert first row to rownames
#'   wb_to_df(wb1, sheet = 3, dims = "C6:G9", rowNames = TRUE)
#'   
#'   # define type of the data.frame
#'   wb_to_df(wb1, cols = c(1, 4), types = c("Var1" = 0, "Var3" = 1))
#'   
#'   # start in row 5
#'   wb_to_df(wb1, startRow = 5, colNames = FALSE)
#'   
#'   # na string
#'   wb_to_df(wb1, na.strings = "")
#'   
#'   # read.xlsx(wb1)
#' 
#'   ###########################################################################
#'   # inlinestr
#'   xlsxFile <- system.file("extdata", "inlinestr.xlsx", package = "openxlsx")
#'   wb2 <- loadWorkbook(xlsxFile)
#'   
#'   # read dataset with inlinestr
#'   wb_to_df(wb2)
#'   # read.xlsx(wb2)
#'   
#'   ###########################################################################
#'   # definedName // namedRegion
#'   xlsxFile <- system.file("extdata", "namedRegions3.xlsx", package = "openxlsx")
#'   wb3 <- loadWorkbook(xlsxFile)
#'  
#'   # read dataset with definedName (returns global first)
#'   wb_to_df(wb3, definedName = "MyRange", colNames = F)
#'   
#'   # read definedName from sheet
#'   wb_to_df(wb3, definedName = "MyRange", sheet = 4, colNames = F)
#' 
#' @export
wb_to_df <- function(xlsxFile,
                     sheet, 
                     startRow = 1,
                     colNames = TRUE, 
                     rowNames = FALSE,
                     detectDates = TRUE,
                     skipEmptyCols = FALSE,
                     skipEmptyRows = FALSE,
                     rows = NULL,
                     cols = NULL,
                     na.strings = "#N/A",
                     dims, 
                     showFormula = FALSE,
                     convert = TRUE,
                     types,
                     definedName) {
  
  if (is.character(xlsxFile)){
    # if using it this way, it might be benefitial to only load the sheet we
    # want to read instead of every sheet of the entire xlsx file WHEN we do
    # not even see it
    wb <- loadWorkbook(xlsxFile)
  } else {
    wb <- xlsxFile
  }
  
  
  if (!missing(definedName)) {
    
    dn <- wb$workbook$definedNames
    wo <- unlist(lapply(dn, function(x) getXML1val(x, "definedName")))
    wo <- gsub("\\$", "", wo)
    wo <- unlist(sapply(wo, strsplit, "!"))
    
    nr <- as.data.frame(
      matrix(wo, 
             ncol = 2, 
             byrow = TRUE, 
             dimnames = list(seq_len(length(dn)),
                             c("sheet", "dims") ))
    )
    nr$name <- sapply(dn, function(x) getXML1attr_one(x, "definedName", "name"))
    nr$local <- sapply(dn, function(x) ifelse(
      openxlsx:::getXML1attr_one(x,"definedName", "localSheetId") == "", 0, 1)
    )
    nr$sheet <- which(wb$sheet_names %in% nr$sheet)
    
    nr <- nr[order(nr$local, nr$name, nr$sheet),]
    
    if (definedName %in% nr$name & missing(sheet)) {
      sel   <- nr[nr$name == definedName, ][1,]
      sheet <- sel$sheet
      dims  <- sel$dims
    } else if (definedName %in% nr$name) {
      sel <- nr[nr$name == definedName & nr$sheet == sheet, ]
      if (NROW(sel) == 0) {
        stop("no such definedName on selected sheet")
      } else {
        dims <- sel$dims
      }
    } else {
      stop("no such definedName")
    }
  }
  
  if (missing(sheet)) sheet <- 1
  
  if (is.character(sheet))
    sheet <- which(wb$sheet_names %in% sheet)
  
  # must be available
  if (missing(dims))
    dims <- getXML1attr_one(wb$worksheets[[sheet]]$dimension,
                            "dimension",
                            "ref")
  
  cc  <- wb$worksheets[[sheet]]$sheet_data$cc
  
  rnams <- names(cc)
  
  
  # internet says: numFmtId > 0 and applyNumberFormat == 1
  sd <- as.data.frame(
    do.call(
      "rbind",
      lapply(
        wb$styles$cellXfs, 
        FUN= function(x) 
          c(
            as.numeric(openxlsx:::getXML1attr_one(x, "xf", "numFmtId")),
            as.numeric(openxlsx:::getXML1attr_one(x, "xf", "applyNumberFormat"))
          )
      ) 
    )
  )
  names(sd) <- c("numFmtId", "applyNumberFormat")
  
  sd$id <- seq_len(nrow(sd))-1
  sd$isdate <- 0
  sd$isdate[sd$numFmtId > 0 &
              sd$applyNumberFormat == 1] <- 1
  
  xlsx_date_style <- sd$id[sd$isdate == 1]
  
  # create temporary data frame
  z <- tt <- dims_to_dataframe(dims)
  
  keep_cols <- colnames(z)
  keep_rows <- rownames(z)
  
  if (startRow > 1) {
    keep_rows <- as.character(seq(startRow, max(as.numeric(keep_rows))))
    
    z  <- z[rownames(z) %in% keep_rows,]
    tt <- tt[rownames(tt) %in% keep_rows,]
  }
  
  if (!is.null(rows)) {
    keep_rows <- as.character(rows)
    
    z  <- z[rownames(z) %in% keep_rows,]
    tt <- tt[rownames(tt) %in% keep_rows,]
  }
  
  if (!is.null(cols)) {
    keep_cols <- int2col(cols)
    
    z  <- z[keep_cols]
    tt <- tt[keep_cols]
  }
  
  keep_row <- keep_rows[keep_rows %in% rnams]
  
  
  for (row in keep_row) {
    
    rowvals    <-  cc[[row]]
    
    keep_col <- keep_cols[keep_cols %in% names(rowvals)]
    
    for (col in keep_col) {
      
      val       <- rowvals[[col]][["val"]]
      
      if (!is.null(val$v) | !is.null(val$is)) {
        
        this_ttyp <- rowvals[[col]][["typ"]][["t"]]
        this_styp <- rowvals[[col]][["typ"]][["s"]]
        
        if (!is.null(this_ttyp)) {
          
          # sharedString: string
          if (this_ttyp == "s") {
            val$v <- wb$sharedStrings[as.numeric(val$v)+1]
            
            tt[[col]][rownames(tt) == row]  <- "s"
          }
          
          # str: should be from function evaluation?
          if (this_ttyp == "str") {
            tt[[col]][rownames(tt) == row]  <- "s"
          }
          
          # inlinestr: string
          if (this_ttyp == "inlineStr") {
            val$v <- val$is
            
            tt[[col]][rownames(tt) == row]  <- "s"
          }
          
          # bool: logical value
          if (this_ttyp == "b") {
            val$v <- as.logical(as.numeric(val$v))
            
            tt[[col]][rownames(tt) == row]  <- "b"
          }
          
          # evaluation: takes the formula value?
          if (showFormula) {
            if(!is.null(val$f)) val$v <- val$f
            
            tt[[col]][rownames(tt) == row]  <- "s"
          }
          
          # convert na.string to NA
          if (!is.na(na.strings) | !missing(na.strings)) {
            if(val$v %in% na.strings) {
              val$v <- NA
              tt[[col]][rownames(tt) == row]  <- NA
            }
          }
        }
        
        # dates
        if (!is.null(this_styp)) {
          
          if (detectDates) {
            if ( (this_styp %in% xlsx_date_style) ) {
              val$v <- as.character(convertToDate(as.numeric(val$v)))
              
              tt[[col]][rownames(tt) == row]  <- "d"
            }
          }
          
        }
        
        # check if val is some kind of string expression
        if ( !is.na(val$v) &  !(tt[[col]][rownames(tt) == row] %in% c("b", "d", "s", "str")) ) {
          
          # check if it becomes NA when changing from character, to numeric and back
          if (suppressWarnings(is.na(as.character(as.numeric(val$v)))))
            tt[[col]][rownames(tt) == row]  <- "s"
        }
        
        z[[col]][rownames(z) == row] <- val$v
        
      }
    }
    
  } # end row loop
  
  # if colNames, then change tt too
  if (colNames) {
    colnames(z)  <- z[1,]
    colnames(tt) <- z[1,]
    
    z  <- z[-1, ]
    tt <- tt[-1, ]
  }
  
  if (rowNames) {
    rownames(z)  <- z[,1]
    rownames(tt) <- z[,1]
    
    z  <- z[ ,-1]
    tt <- tt[ , -1]
  }
  
  if (skipEmptyRows) {
    empty <- apply(z, 1, function(x) all(is.na(x)), simplify = TRUE)
    
    z  <- z[!empty, ]
    tt <- tt[!empty,]
  }
  
  if (skipEmptyCols) {
    
    empty <- sapply(z, function(x) all(is.na(x)))
    
    if (any(empty)) {
      sel <- names(empty[empty == TRUE])
      z[sel]  <- NULL
      tt[sel] <- NULL
    }
    
  }
  
  if (missing(types))
    types <- guess_col_type(tt)
  
  # could make it optional or explicit
  if (convert) {
    nums <- names(types[types == 1])
    dtes <- names(types[types == 2])
    
    z[nums] <- lapply(z[nums], as.numeric)
    z[dtes] <- lapply(z[dtes], as.Date)
  }
  
  attr(z, "tt") <- tt
  attr(z, "types") <- types
  if (!missing(definedName)) attr(z, "dn") <- nr
  z
}
