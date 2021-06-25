
#' create dataframe from dimensions
#' @param dims Character vector of expected dimension.
#' @examples {
#'   dims_to_dataframe("A1:B2")
#' }
#' @export
dims_to_dataframe <- function(dims) {
  
  dimensions <- strsplit(dims, ":")
  rows <- as.numeric(gsub("[[:upper:]]","", dimensions[[1]]))
  cols <- gsub("[[:digit:]]","", dimensions[[1]])
  
  # cols
  col_min <- openxlsx:::cell_ref_to_col(cols[1])
  col_max <- openxlsx:::cell_ref_to_col(cols[2])
  
  cols <- seq(col_min, col_max)
  cols <- openxlsx::int2col(cols)
  
  # rows
  rows_min <- rows[1]
  rows_max <- rows[2]
  
  rows <- seq(rows_min, rows_max)
  
  # matrix as.data.frame
  mm <- matrix(data = NA,
               nrow = length(rows),
               ncol = length(cols),
               dimnames = list(rows, cols))
  
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
#' @param wb openxlsx workbook
#' @param sheet Either sheet name or index. When missing the first sheet in the workbook is selected.
#' @param colnames If TRUE, the first row of data will be used as column names.
#' @param dims Character string of type "A1:B2" as optional dimentions to be imported.
#' @param detectDates If TRUE, attempt to recognise dates and perform conversion.
#' @param showFormula If TRUE, the underlying Excel formulas are shown.
#' @param convert If TRUE, a conversion to dates and numerics is attempted.
#' @param skipEmptyCols If TRUE, empty columns are skipped.
#' @examples
#'   # numerics, dates, missings, bool and string
#'   xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx")
#'   wb1 <- loadWorkbook(xlsxFile)
#'   
#'   # inlinestr
#'   xlsxFile <- system.file("extdata", "inlinestr.xlsx", package = "openxlsx")
#'   wb2 <- loadWorkbook(xlsxFile)
#' 
#'   wb_to_df(wb1)
#'   wb_to_df(wb1, colNames = FALSE)
#'   wb_to_df(wb1, detectDates = FALSE)
#'   wb_to_df(wb1, showFormula = TRUE)
#'   wb_to_df(wb1, dims = "A2:C5", colNames = FALSE)
#'   wb_to_df(wb1, convert = FALSE)
#'   wb_to_df(wb1, skipEmptyCols = TRUE)
#'   read.xlsx(wb1)
#' 
#'   wb_to_df(wb2)
#'   read.xlsx(wb2)
#' 
#' @export
wb_to_df <- function(wb, sheet, colNames = TRUE, dims, detectDates = TRUE,
                     showFormula = FALSE, convert = TRUE,
                     skipEmptyCols = FALSE) {
  
  if (missing(sheet)) sheet <- 1
  
  if (is.character(sheet))
    sheet <- which(wb1$sheet_names %in% sheet)
  
  # must be available
  if (missing(dims))
    dims <- getXML1attr_one(wb$worksheets[[sheet]]$dimension,
                            "dimension",
                            "ref")
  
  vval  <- wb$worksheets[[sheet]]$sheet_data$vval
  fval  <- wb$worksheets[[sheet]]$sheet_data$fval
  ttyp  <- wb$worksheets[[sheet]]$sheet_data$ttyp  # to identify strings and numbers
  styp  <- wb$worksheets[[sheet]]$sheet_data$styp  # to identify dates?
  rtyp  <- wb$worksheets[[sheet]]$sheet_data$rtyp  # to identify dates?
  isval <- wb$worksheets[[sheet]]$sheet_data$isval # inlinestr
  
  rnams <- get_row_names(rtyp)
  names(vval)  <- rnams
  names(fval)  <- rnams
  names(ttyp)  <- rnams
  names(styp)  <- rnams
  names(rtyp)  <- rnams
  names(isval) <- rnams
  
  
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
  
  keep_col <- colnames(z)
  keep_row <- rownames(z)
  
  for (row in keep_row) {
    
    if ((row %in% keep_row)  & (row %in% rnams)) {
      
      rowvals   <- vval[[row]]
      rowvals_f <- fval[[row]]
      
      for (col in seq_along(rowvals)) {
        nam <- names(rowvals[col])
        
        if (nam %in% keep_col) {
          
          val <- unlist(rowvals[col])
          
          if (ttyp[[row]][col] == "s") {
            val <- wb$sharedStrings[as.numeric(val)+1]
            val <- getXML2val(val, level1 = "si", child = "t")
            
            tt[[nam]][rownames(tt) == row]  <- "s"
          }
          
          if (ttyp[[row]][col] == "str") {
            tt[[nam]][rownames(tt) == row]  <- "s"
          }
          
          if (ttyp[[row]][col] == "b") {
            val <- as.logical(as.numeric(val))
            
            tt[[nam]][rownames(tt) == row]  <- "b"
          }
          
          if (ttyp[[row]][col] == "e") {
            if (showFormula) {
              tmp <- unlist(rowvals_f[col])
              if (!identical(tmp, character(0))) {
                val <- tmp
                
                tt[[nam]][rownames(tt) == row]  <- "s"
              }
            }
          }
          
          if (detectDates) {
            if (styp[[row]][col] %in% xlsx_date_style ) {
              val <- as.character(convertToDate(as.numeric(val)))
              
              tt[[nam]][rownames(tt) == row]  <- "d"
            }
          }
          
          if (!identical(val, character(0))) {
            
            # check if val is some kind of string expression
            if ( !(tt[[nam]][rownames(tt) == row] %in% c("b", "d")) )
              if (suppressWarnings(is.na(as.character(as.numeric(val)))))
                tt[[nam]][rownames(tt) == row]  <- "s"
            
            z[[nam]][rownames(z) == row] <- val
          }
          
        }
      }
    }
  }
  
  # overwrite for 
  for (row in keep_row) {
    
    if ((row %in% keep_row) & (row %in% rnams)) {
      rowvals   <- isval[[row]]
      
      for (col in seq_along(rowvals)) {
        nam <- names(rowvals[col])
        
        if (nam %in% keep_col) {
          
          val <- unlist(rowvals[col])
          
          
          # keep only non character(0)
          if (!identical(val, character(0))) {
            z[[nam]][rownames(z) == row] <- val
            
            tt[[nam]][rownames(tt) == row]  <- "s"
          }
        }
      }
    }
  }
  
  # if colNames, then change tt too
  if (colNames) {
    colnames(z)  <- z[1,]
    colnames(tt) <- z[1,]
    
    z  <- z[-1, ]
    tt <- tt[-1, ]
  }
  
  types <- guess_col_type(tt)
  
  # could make it optional or explicit
  if (convert) {
    nums <- names(types[types == 1])
    dtes <- names(types[types == 2])
    
    z[nums] <- lapply(z[nums], as.numeric)
    z[dtes] <- lapply(z[dtes], as.Date)
  }
  
  if (skipEmptyCols) {
    
    empty <- sapply(z, function(x) all(is.na(x)))
    
    if (any(empty)) {
      sel <- names(empty[empty == TRUE])
      z[sel]  <- NULL
      tt[sel] <- NULL
    }
    
  }
  
  attr(z, "tt") <- tt
  attr(z, "types") <- types
  z
}
