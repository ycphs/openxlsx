
#' @name writeDataTable
#' @title Write to a worksheet as an Excel table
#' @description Write to a worksheet and format as an Excel table
#' @param wb A Workbook object containing a
#' worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x A dataframe.
#' @param startCol A vector specifying the starting column to write df
#' @param startRow A vector specifying the starting row to write df
#' @param xy An alternative to specifying startCol and startRow individually.
#' A vector of the form c(startCol, startRow)
#' @param colNames If `TRUE`, column names of x are written.
#' @param rowNames If `TRUE`, row names of x are written.
#' @param row.names,col.names Deprecated, please use `rowNames`, `colNames` instead
#' @param tableStyle Any excel table style name or "none" (see "formatting" vignette).
#' @param tableName name of table in workbook. The table name must be unique.
#' @param headerStyle Custom style to apply to column names.
#' @param withFilter If `TRUE` or `NA`, columns with have filters in the first row.
#' @param keepNA If `TRUE`, NA values are converted to #N/A (or `na.string`, if not NULL) in Excel, else NA cells will be empty.
#' @param na.string If not NULL, and if `keepNA` is `TRUE`, NA values are converted to this string in Excel.
#' @param sep Only applies to list columns. The separator used to collapse list columns to a character vector e.g. sapply(x$list_column, paste, collapse = sep).
#' @param stack If `TRUE` the new style is merged with any existing cell styles.  If FALSE, any
#' existing style is replaced by the new style.
#' \cr\cr
#' \cr**The below options correspond to Excel table options:**
#' \cr
#' \if{html}{\figure{tableoptions.png}{options: width="40\%" alt="Figure: table_options.png"}}
#' \if{latex}{\figure{tableoptions.pdf}{options: width=7cm}}
#'
#' @param firstColumn logical. If TRUE, the first column is bold
#' @param lastColumn logical. If TRUE, the last column is bold
#' @param bandedRows logical. If TRUE, rows are colour banded
#' @param bandedCols logical. If TRUE, the columns are colour banded
#' @details columns of x with class Date/POSIXt, currency, accounting,
#' hyperlink, percentage are automatically styled as dates, currency, accounting,
#' hyperlinks, percentages respectively.
#' @seealso [addWorksheet()]
#' @seealso [writeData()]
#' @seealso [removeTable()]
#' @seealso [getTables()]
#' @export
#' @examples
#' ## see package vignettes for further examples.
#'
#' #####################################################################################
#' ## Create Workbook object and add worksheets
#' wb <- createWorkbook()
#' addWorksheet(wb, "S1")
#' addWorksheet(wb, "S2")
#' addWorksheet(wb, "S3")
#'
#'
#' #####################################################################################
#' ## -- write data.frame as an Excel table with column filters
#' ## -- default table style is "TableStyleMedium2"
#'
#' writeDataTable(wb, "S1", x = iris)
#'
#' writeDataTable(wb, "S2",
#'   x = mtcars, xy = c("B", 3), rowNames = TRUE,
#'   tableStyle = "TableStyleLight9"
#' )
#'
#' df <- data.frame(
#'   "Date" = Sys.Date() - 0:19,
#'   "T" = TRUE, "F" = FALSE,
#'   "Time" = Sys.time() - 0:19 * 60 * 60,
#'   "Cash" = paste("$", 1:20), "Cash2" = 31:50,
#'   "hLink" = "https://CRAN.R-project.org/",
#'   "Percentage" = seq(0, 1, length.out = 20),
#'   "TinyNumbers" = runif(20) / 1E9, stringsAsFactors = FALSE
#' )
#'
#' ## openxlsx will apply default Excel styling for these classes
#' class(df$Cash) <- c(class(df$Cash), "currency")
#' class(df$Cash2) <- c(class(df$Cash2), "accounting")
#' class(df$hLink) <- "hyperlink"
#' class(df$Percentage) <- c(class(df$Percentage), "percentage")
#' class(df$TinyNumbers) <- c(class(df$TinyNumbers), "scientific")
#'
#' writeDataTable(wb, "S3", x = df, startRow = 4, rowNames = TRUE, tableStyle = "TableStyleMedium9")
#'
#' #####################################################################################
#' ## Additional Header Styling and remove column filters
#'
#' writeDataTable(wb,
#'   sheet = 1, x = iris, startCol = 7, headerStyle = createStyle(textRotation = 45),
#'   withFilter = FALSE
#' )
#'
#'
#' #####################################################################################
#' ## Save workbook
#' ## Open in excel without saving file: openXL(wb)
#' \dontrun{
#' saveWorkbook(wb, "writeDataTableExample.xlsx", overwrite = TRUE)
#' }
#'
#'
#'
#'
#'
#' #####################################################################################
#' ## Pre-defined table styles gallery
#'
#' wb <- createWorkbook(paste0("tableStylesGallery.xlsx"))
#' addWorksheet(wb, "Style Samples")
#' for (i in 1:21) {
#'   style <- paste0("TableStyleLight", i)
#'   writeDataTable(wb,
#'     x = data.frame(style), sheet = 1,
#'     tableStyle = style, startRow = 1, startCol = i * 3 - 2
#'   )
#' }
#'
#' for (i in 1:28) {
#'   style <- paste0("TableStyleMedium", i)
#'   writeDataTable(wb,
#'     x = data.frame(style), sheet = 1,
#'     tableStyle = style, startRow = 4, startCol = i * 3 - 2
#'   )
#' }
#'
#' for (i in 1:11) {
#'   style <- paste0("TableStyleDark", i)
#'   writeDataTable(wb,
#'     x = data.frame(style), sheet = 1,
#'     tableStyle = style, startRow = 7, startCol = i * 3 - 2
#'   )
#' }
#'
#' ## openXL(wb)
#' \dontrun{
#' saveWorkbook(wb, file = "tableStylesGallery.xlsx", overwrite = TRUE)
#' }
#'
writeDataTable <- function(
  wb,
  sheet,
  x,
  startCol    = 1,
  startRow    = 1,
  xy          = NULL,
  colNames    = TRUE,
  rowNames    = FALSE,
  tableStyle  = openxlsx_getOp("tableStyle", "TableStyleLight9"),
  tableName   = NULL,
  headerStyle = openxlsx_getOp("headerStyle"),
  withFilter  = openxlsx_getOp("withFilter", TRUE),
  keepNA      = openxlsx_getOp("keepNA", FALSE),
  na.string   = openxlsx_getOp("na.string"),
  sep         = ", ",
  stack       = FALSE,
  firstColumn = openxlsx_getOp("firstColumn", FALSE),
  lastColumn  = openxlsx_getOp("lastColumn", FALSE),
  bandedRows  = openxlsx_getOp("bandedRows", TRUE),
  bandedCols  = openxlsx_getOp("bandedCols", FALSE),
  col.names,
  row.names
  ) {
  op <- get_set_options()
  on.exit(options(op), add = TRUE)
  
  ## increase scipen to avoid writing in scientific
  
  if (!missing(row.names)) {
    warning("Please use 'rowNames' instead of 'row.names'", call. = FALSE)
    row.names <- rowNames
  }
  
  if (!missing(col.names)) {
    warning("Please use 'colNames' instead of 'col.names'", call. = FALSE)
    colNames <- col.names
  }
  
  # Set NULLs
  withFilter  <- withFilter  %||% TRUE
  keepNA      <- keepNA      %||% FALSE
  firstColumn <- firstColumn %||% FALSE
  lastColumn  <- lastColumn  %||% FALSE
  bandedRows  <- bandedRows  %||% TRUE
  bandedCols  <- bandedCols  %||% FALSE
  withFilter  <- withFilter  %||% TRUE
  
  if (!is.null(xy)) {
    if (length(xy) != 2) {
      stop("xy parameter must have length 2")
    }
    startCol <- xy[[1]]
    startRow <- xy[[2]]
  }

  # Assert parameters
  assert_class(wb, "Workbook")
  assert_class(x, "data.frame")
  assert_true_false(colNames)
  assert_true_false(rowNames)
  assert_class(headerStyle, "Style", or_null = TRUE)
  assert_true_false(withFilter)
  assert_character1(sep)
  assert_true_false(firstColumn)
  assert_true_false(lastColumn)
  assert_true_false(bandedRows)
  assert_true_false(bandedCols)
  
  if (is.null(tableName)) {
    tableName <- sprintf("Table%i", length(wb$tables) + 3L)
  } else {
    tableName <- wb$validate_table_name(tableName)
  }

  ## convert startRow and startCol
  if (!is.numeric(startCol)) {
    startCol <- convertFromExcelRef(startCol)
  }
  startRow <- as.integer(startRow)

  ## Coordinates for each section
  if (rowNames) {
    x <- cbind(data.frame("row names" = rownames(x)), as.data.frame(x))
  }

  ## If 0 rows append a blank row
  
  tableStyle <- validate_StyleName(tableStyle)

  ## header style
  if (inherits(headerStyle, "Style")) {
    addStyle(
      wb         = wb, 
      sheet      = sheet,
      style      = headerStyle,
      rows       = startRow,
      cols       = 0:(ncol(x) - 1L) + startCol,
      gridExpand = TRUE
    )
  }

  showColNames <- colNames

  if (colNames) {
    colNames <- colnames(x)
    assert_unique(colNames, case_sensitive = FALSE)

    ## zero char names are invalid
    char0 <- nchar(colNames) == 0
    if (any(char0)) {
      colNames[char0] <- colnames(x)[char0] <- paste0("Column", which(char0))
    }
    
    # Compatibility with MS Excel: throw warning if a table column name exceeds
    # the length of 255 chars.
    char_over255 <- nchar(colNames) > 255
    if (any(char_over255)) {
      warning_msg <- sprintf(
        "Column name exceeds 255 chars, possible incompatibility with MS Excel. Index: %s",
        toString(which(char_over255)))
      warning(warning_msg)
    }

  } else {
    colNames <- paste0("Column", seq_along(x))
    names(x) <- colNames
  }
  
  ## If zero rows, append an empty row (prevent XML from corrupting)
  if (nrow(x) == 0) {
    x <- rbind(
      as.data.frame(x),
      matrix("", nrow = 1, ncol = ncol(x), dimnames = list(character(), colnames(x)))
    )
    names(x) <- colNames
  }

  ref1 <- paste0(convert_to_excel_ref(cols = startCol, LETTERS = LETTERS), startRow)
  ref2 <- paste0(convert_to_excel_ref(cols = startCol + ncol(x) - 1, LETTERS = LETTERS), startRow + nrow(x))
  ref <- paste(ref1, ref2, sep = ":")

  ## check not overwriting another table
  wb$check_overwrite_tables(
    sheet = sheet,
    new_rows = c(startRow, startRow + nrow(x) - 1L + 1L), ## + header
    new_cols = c(startCol, startCol + ncol(x) - 1L)
  )


  ## column class styling
  # consider not using lowercase and instead use inherits(x, class)
  colClasses <- lapply(x, function(x) tolower(class(x)))
  classStyles(
    wb,
    sheet      = sheet,
    startRow   = startRow, 
    startCol   = startCol, 
    colNames   = TRUE,
    nRow       = nrow(x),
    colClasses = colClasses, 
    stack      = stack
  )

  ## write data to worksheet
  wb$writeData(
    df         = x,
    colNames   = TRUE,
    sheet      = sheet,
    startRow   = startRow,
    startCol   = startCol,
    colClasses = colClasses,
    hlinkNames = NULL,
    keepNA     = keepNA,
    na.string  = na.string,
    list_sep   = sep
  )

  ## replace invalid XML characters
  colNames <- replaceIllegalCharacters(colNames)

  ## create table.xml and assign an id to worksheet tables
  wb$buildTable(
    sheet             = sheet,
    colNames          = colNames,
    ref               = ref,
    showColNames      = showColNames,
    tableStyle        = tableStyle,
    tableName         = tableName,
    withFilter        = withFilter[1],
    totalsRowCount    = 0L,
    showFirstColumn   = firstColumn[1],
    showLastColumn    = lastColumn[1],
    showRowStripes    = bandedRows[1],
    showColumnStripes = bandedCols[1]
  )
}
