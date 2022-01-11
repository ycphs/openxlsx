#' @name writeData
#' @title Write an object to a worksheet
#' @author Alexander Walker
#' @import stringi
#' @description Write an object to worksheet with optional styling.
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x Object to be written. For classes supported look at the examples.
#' @param startCol A vector specifying the starting column to write to.
#' @param startRow A vector specifying the starting row to write to.
#' @param array A bool if the function written is of type array
#' @param xy An alternative to specifying `startCol` and
#' `startRow` individually.  A vector of the form
#' `c(startCol, startRow)`.
#' @param colNames If `TRUE`, column names of x are written.
#' @param rowNames If `TRUE`, data.frame row names of x are written.
#' @param row.names,col.names Deprecated, please use `rowNames`, `colNames` instead
#' @param headerStyle Custom style to apply to column names.
#' @param borders Either "`none`" (default), "`surrounding`",
#' "`columns`", "`rows`" or *respective abbreviations*.  If
#' "`surrounding`", a border is drawn around the data.  If "`rows`",
#' a surrounding border is drawn with a border around each row. If
#' "`columns`", a surrounding border is drawn with a border between
#' each column. If "`all`" all cell borders are drawn.
#' @param borderColour Colour of cell border.  A valid colour (belonging to `colours()` or a hex colour code, eg see [here](https://www.w3schools.com/web-design/color-picker/)).
#' @param borderStyle Border line style
#' \itemize{
#'    \item{**none**}{ no border}
#'    \item{**thin**}{ thin border}
#'    \item{**medium**}{ medium border}
#'    \item{**dashed**}{ dashed border}
#'    \item{**dotted**}{ dotted border}
#'    \item{**thick**}{ thick border}
#'    \item{**double**}{ double line border}
#'    \item{**hair**}{ hairline border}
#'    \item{**mediumDashed**}{ medium weight dashed border}
#'    \item{**dashDot**}{ dash-dot border}
#'    \item{**mediumDashDot**}{ medium weight dash-dot border}
#'    \item{**dashDotDot**}{ dash-dot-dot border}
#'    \item{**mediumDashDotDot**}{ medium weight dash-dot-dot border}
#'    \item{**slantDashDot**}{ slanted dash-dot border}
#'   }
#' @param withFilter If `TRUE` or `NA`, add filters to the column name row. NOTE can only have one filter per worksheet.
#' @param keepNA If `TRUE`, NA values are converted to #N/A (or `na.string`, if not NULL) in Excel, else NA cells will be empty.
#' @param na.string If not NULL, and if `keepNA` is `TRUE`, NA values are converted to this string in Excel.
#' @param name If not NULL, a named region is defined.
#' @param sep Only applies to list columns. The separator used to collapse list columns to a character vector e.g. sapply(x$list_column, paste, collapse = sep).
#' @seealso [writeDataTable()]
#' @export writeData
#' @details Formulae written using writeFormula to a Workbook object will not get picked up by read.xlsx().
#' This is because only the formula is written and left to Excel to evaluate the formula when the file is opened in Excel.
#' @rdname writeData
#' @return invisible(0)
#' @examples
#'
#' ## See formatting vignette for further examples.
#'
#' ## Options for default styling (These are the defaults)
#' options("openxlsx.borderColour" = "black")
#' options("openxlsx.borderStyle" = "thin")
#' options("openxlsx.dateFormat" = "mm/dd/yyyy")
#' options("openxlsx.datetimeFormat" = "yyyy-mm-dd hh:mm:ss")
#' options("openxlsx.numFmt" = NULL)
#'
#' ## Change the default border colour to #4F81BD
#' options("openxlsx.borderColour" = "#4F81BD")
#'
#'
#' #####################################################################################
#' ## Create Workbook object and add worksheets
#' wb <- createWorkbook()
#'
#' ## Add worksheets
#' addWorksheet(wb, "Cars")
#' addWorksheet(wb, "Formula")
#'
#'
#' x <- mtcars[1:6, ]
#' writeData(wb, "Cars", x, startCol = 2, startRow = 3, rowNames = TRUE)
#'
#' #####################################################################################
#' ## Bordering
#'
#' writeData(wb, "Cars", x,
#'   rowNames = TRUE, startCol = "O", startRow = 3,
#'   borders = "surrounding", borderColour = "black"
#' ) ## black border
#'
#' writeData(wb, "Cars", x,
#'   rowNames = TRUE,
#'   startCol = 2, startRow = 12, borders = "columns"
#' )
#'
#' writeData(wb, "Cars", x,
#'   rowNames = TRUE,
#'   startCol = "O", startRow = 12, borders = "rows"
#' )
#'
#'
#' #####################################################################################
#' ## Header Styles
#'
#' hs1 <- createStyle(
#'   fgFill = "#DCE6F1", halign = "CENTER", textDecoration = "italic",
#'   border = "Bottom"
#' )
#'
#' writeData(wb, "Cars", x,
#'   colNames = TRUE, rowNames = TRUE, startCol = "B",
#'   startRow = 23, borders = "rows", headerStyle = hs1, borderStyle = "dashed"
#' )
#'
#'
#' hs2 <- createStyle(
#'   fontColour = "#ffffff", fgFill = "#4F80BD",
#'   halign = "center", valign = "center", textDecoration = "bold",
#'   border = "TopBottomLeftRight"
#' )
#'
#' writeData(wb, "Cars", x,
#'   colNames = TRUE, rowNames = TRUE,
#'   startCol = "O", startRow = 23, borders = "columns", headerStyle = hs2
#' )
#'
#'
#'
#'
#' #####################################################################################
#' ## Hyperlinks
#' ## - vectors/columns with class 'hyperlink' are written as hyperlinks'
#'
#' v <- rep("https://CRAN.R-project.org/", 4)
#' names(v) <- paste0("Hyperlink", 1:4) # Optional: names will be used as display text
#' class(v) <- "hyperlink"
#' writeData(wb, "Cars", x = v, xy = c("B", 32))
#'
#'
#' #####################################################################################
#' ## Formulas
#' ## - vectors/columns with class 'formula' are written as formulas'
#'
#' df <- data.frame(
#'   x = 1:3, y = 1:3,
#'   z = paste0(paste0("A", 1:3 + 1L), paste0("B", 1:3 + 1L), sep = " + "),
#'   stringsAsFactors = FALSE
#' )
#'
#' class(df$z) <- c(class(df$z), "formula")
#'
#' writeData(wb, sheet = "Formula", x = df)
#'
#'
#' #####################################################################################
#' ## Save workbook
#' ## Open in excel without saving file: openXL(wb)
#' \dontrun{
#' saveWorkbook(wb, "writeDataExample.xlsx", overwrite = TRUE)
#' }
writeData <- function(
  wb,
  sheet,
  x,
  startCol     = 1,
  startRow     = 1,
  array        = FALSE,
  xy           = NULL,
  colNames     = TRUE,
  rowNames     = FALSE,
  headerStyle  = openxlsx_getOp("headerStyle"),
  borders      = openxlsx_getOp("borders", "none"),
  borderColour = openxlsx_getOp("borderColour", "black"),
  borderStyle  = openxlsx_getOp("borderStyle", "thin"),
  withFilter   = openxlsx_getOp("withFilter", FALSE),
  keepNA       = openxlsx_getOp("keepNA", FALSE),
  na.string    = openxlsx_getOp("na.string"),
  name         = NULL,
  sep          = ", ",
  col.names,
  row.names
) {

  x <- force(x)
  
  op <- get_set_options()
  on.exit(options(op), add = TRUE)
  
  if (!missing(row.names)) {
    warning("Please use 'rowNames' instead of 'row.names'", call. = FALSE)
    rowNames <- row.names
  }
  
  if (!missing(col.names)) {
    warning("Please use 'colNames' instead of 'col.names'", call. = FALSE)
    colNames <- col.names
  }
  
  # Set NULLs
  borders      <- borders      %||% "none"
  borderColour <- borderColour %||% "black"
  borderStyle  <- borderStyle  %||% "thin"
  withFilter   <- withFilter   %||% FALSE
  keepNA       <- keepNA       %||% FALSE

  if (is.null(x)) {
    return(invisible(0))
  }

  ## All input conversions/validations
  if (!is.null(xy)) {
    if (length(xy) != 2) {
      stop("xy parameter must have length 2")
    }
    startCol <- xy[[1]]
    startRow <- xy[[2]]
  }

  ## convert startRow and startCol
  if (!is.numeric(startCol)) {
    startCol <- convertFromExcelRef(startCol)
  }

  startRow <- as.integer(startRow)
  
  assert_class(wb, "Workbook")
  assert_true_false(colNames)
  assert_true_false(rowNames)
  assert_character1(sep)
  assert_class(headerStyle, "Style", or_null = TRUE)

  ## borderColours validation
  borderColour <- validateColour(borderColour, "Invalid border colour")
  borderStyle <- validateBorderStyle(borderStyle)[[1]]

  ## special case - vector of hyperlinks
  hlinkNames <- NULL
  if (inherits(x, "hyperlink")) {
    hlinkNames <- names(x)
    colNames <- FALSE
  }

  ## special case - formula
  if (inherits(x, "formula")) {
    x <- data.frame("X" = x, stringsAsFactors = FALSE)
    class(x[[1]]) <- ifelse(array, "array_formula", "formula")
    colNames <- FALSE
  }

  ## named region
  if (!is.null(name)) { ## validate name
    ex_names <- regmatches(wb$workbook$definedNames, regexpr('(?<=name=")[^"]+', wb$workbook$definedNames, perl = TRUE))
    ex_names <- replaceXMLEntities(ex_names)

    if (name %in% ex_names) {
      stop(sprintf("Named region with name '%s' already exists!", name))
    } else if (grepl("^[A-Z]{1,3}[0-9]+$", name)) {
      stop("name cannot look like a cell reference.")
    }
  }

  if (is.vector(x) | is.factor(x) | inherits(x, "Date")) {
    colNames <- FALSE
  } ## this will go to coerce.default and rowNames will be ignored

  ## Coerce to data.frame
  x <- openxlsxCoerce(x = x, rowNames = rowNames)

  nCol <- ncol(x)
  nRow <- nrow(x)

  ## If no rows and not writing column names return as nothing to write
  if (nRow == 0 & !colNames) {
    return(invisible(0))
  }

  ## If no columns and not writing row names return as nothing to write
  if (nCol == 0 & !rowNames) {
    return(invisible(0))
  }

  colClasses <- lapply(x, function(x) tolower(class(x)))
  colClasss2 <- colClasses
  colClasss2[vapply(
    colClasses,
    function(i) inherits(i, "formula") & inherits(i, "hyperlink"),
    NA
  )] <- "formula"

  if (is.numeric(sheet)) {
    sheetX <- wb$validateSheet(sheet)
  } else {
    sheetX <- wb$validateSheet(replaceXMLEntities(sheet))
    sheet <- replaceXMLEntities(sheet)
  }

  if (wb$isChartSheet[[sheetX]]) {
    stop("Cannot write to chart sheet.")
  }

  ## Check not overwriting existing table headers
  wb$check_overwrite_tables(
    sheet                   = sheet,
    new_rows                = c(startRow, startRow + nRow - 1L + colNames),
    new_cols                = c(startCol, startCol + nCol - 1L),
    check_table_header_only = TRUE,
    error_msg               = "Cannot overwrite table headers. Avoid writing over the header row or see getTables() & removeTables() to remove the table object."
  )

  ## write autoFilter, can only have a single filter per worksheet
  if (withFilter) {
    coords <- data.frame(
      x = c(startRow, startRow + nRow + colNames - 1L),
      y = c(startCol, startCol + nCol - 1L)
    )
    
    ref <- stri_join(getCellRefs(coords), collapse = ":")
    wb$worksheets[[sheetX]]$autoFilter <- sprintf('<autoFilter ref="%s"/>', ref)
    l <- convert_to_excel_ref(cols = unlist(coords[, 2]), LETTERS = LETTERS)
    dfn <- sprintf("'%s'!%s", names(wb)[sheetX], stri_join("$", l, "$", coords[, 1], collapse = ":"))

    dn <- sprintf('<definedName name="_xlnm._FilterDatabase" localSheetId="%s" hidden="1">%s</definedName>', sheetX - 1L, dfn)

    if (length(wb$workbook$definedNames) > 0) {
      ind <- grepl('name="_xlnm._FilterDatabase"', wb$workbook$definedNames)
      if (length(ind) > 0) {
        wb$workbook$definedNames[ind] <- dn
      }
    } else {
      wb$workbook$definedNames <- dn
    }
  }

  ## write data.frame
  wb$writeData(
    df         = x,
    colNames   = colNames,
    sheet      = sheet,
    startCol   = startCol,
    startRow   = startRow,
    colClasses = colClasss2,
    hlinkNames = hlinkNames,
    keepNA     = keepNA,
    na.string  = na.string,
    list_sep   = sep
  )

  ## header style
  if (inherits(headerStyle, "Style") & colNames) {
    addStyle(
      wb         = wb,
      sheet      = sheet, 
      style      = headerStyle,
      rows       = startRow,
      cols       = 0:(nCol - 1) + startCol,
      gridExpand = TRUE,
      stack      = TRUE
    )
  }

  ## If we don't have any rows to write return
  if (nRow == 0) {
    return(invisible(0))
  }

  ## named region
  if (!is.null(name)) {
    ref1 <- stri_join("$", convert_to_excel_ref(cols = startCol, LETTERS = LETTERS), "$", startRow)
    ref2 <- stri_join("$", convert_to_excel_ref(cols = startCol + nCol - 1L, LETTERS = LETTERS), "$", startRow + nRow - 1L + colNames)
    wb$createNamedRegion(ref1 = ref1, ref2 = ref2, name = name, sheet = wb$sheet_names[wb$validateSheet(sheet)])
  }

  ## hyperlink style, if no borders
  borders <- match.arg(borders, c("none", "surrounding", "rows", "columns", "all"))

  if (borders == "none") {
    invisible(
      classStyles(
        wb,
        sheet      = sheet,
        startRow   = startRow,
        startCol   = startCol,
        colNames   = colNames,
        nRow       = nrow(x),
        colClasses = colClasses,
        stack      = TRUE
      )
    )
  } else if (borders == "surrounding") {
    wb$surroundingBorders(
      colClasses,
      sheet        = sheet,
      startRow     = startRow + colNames,
      startCol     = startCol,
      nRow         = nRow, nCol = nCol,
      borderColour = list("rgb" = borderColour),
      borderStyle  = borderStyle
    )
  } else if (borders == "rows") {
    wb$rowBorders(
      colClasses,
      sheet        = sheet,
      startRow     = startRow + colNames,
      startCol     = startCol,
      nRow         = nRow, nCol = nCol,
      borderColour = list("rgb" = borderColour),
      borderStyle  = borderStyle
    )
  } else if (borders == "columns") {
    wb$columnBorders(
      colClasses,
      sheet        = sheet,
      startRow     = startRow + colNames,
      startCol     = startCol,
      nRow         = nRow, nCol = nCol,
      borderColour = list("rgb" = borderColour),
      borderStyle  = borderStyle
    )
  } else if (borders == "all") {
    wb$allBorders(
      colClasses,
      sheet        = sheet,
      startRow     = startRow + colNames,
      startCol     = startCol,
      nRow         = nRow, nCol = nCol,
      borderColour = list("rgb" = borderColour),
      borderStyle  = borderStyle
    )
  }

  invisible(0)
}


#' @name writeFormula
#' @title Write a character vector as an Excel Formula
#' @author Alexander Walker
#' @description Write a a character vector containing Excel formula to a worksheet.
#' @details Currently only the english version of functions are supported. Please don't use the local translation.
#' The examples below show a small list of possible formulas:
#' \itemize{
#'     \item{SUM(B2:B4)}
#'     \item{AVERAGE(B2:B4)}
#'     \item{MIN(B2:B4)}
#'     \item{MAX(B2:B4)}
#'     \item{...}
#'
#' }
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x A character vector.
#' @param startCol A vector specifying the starting column to write to.
#' @param startRow A vector specifying the starting row to write to.
#' @param array A bool if the function written is of type array
#' @param xy An alternative to specifying `startCol` and
#' `startRow` individually.  A vector of the form
#' `c(startCol, startRow)`.
#' @seealso [writeData()] [makeHyperlinkString()]
#' @export writeFormula
#' @rdname writeFormula
#' @examples
#'
#' ## There are 3 ways to write a formula
#'
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet 1")
#' writeData(wb, "Sheet 1", x = iris)
#'
#' ## SEE int2col() to convert int to Excel column label
#'
#' ## 1. -  As a character vector using writeFormula
#'
#' v <- c("SUM(A2:A151)", "AVERAGE(B2:B151)") ## skip header row
#' writeFormula(wb, sheet = 1, x = v, startCol = 10, startRow = 2)
#' writeFormula(wb, 1, x = "A2 + B2", startCol = 10, startRow = 10)
#'
#'
#' ## 2. - As a data.frame column with class "formula" using writeData
#'
#' df <- data.frame(
#'   x = 1:3,
#'   y = 1:3,
#'   z = paste(paste0("A", 1:3 + 1L), paste0("B", 1:3 + 1L), sep = " + "),
#'   z2 = sprintf("ADDRESS(1,%s)", 1:3),
#'   stringsAsFactors = FALSE
#' )
#'
#' class(df$z) <- c(class(df$z), "formula")
#' class(df$z2) <- c(class(df$z2), "formula")
#'
#' addWorksheet(wb, "Sheet 2")
#' writeData(wb, sheet = 2, x = df)
#'
#'
#'
#' ## 3. - As a vector with class "formula" using writeData
#'
#' v2 <- c("SUM(A2:A4)", "AVERAGE(B2:B4)", "MEDIAN(C2:C4)")
#' class(v2) <- c(class(v2), "formula")
#'
#' writeData(wb, sheet = 2, x = v2, startCol = 10, startRow = 2)
#'
#' ## Save workbook
#' \dontrun{
#' saveWorkbook(wb, "writeFormulaExample.xlsx", overwrite = TRUE)
#' }
#'
#'
#' ## 4. - Writing internal hyperlinks
#'
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet1")
#' addWorksheet(wb, "Sheet2")
#' writeFormula(wb, "Sheet1", x = '=HYPERLINK("#Sheet2!B3", "Text to Display - Link to Sheet2")')
#'
#' ## Save workbook
#' \dontrun{
#' saveWorkbook(wb, "writeFormulaHyperlinkExample.xlsx", overwrite = TRUE)
#' }
#'
writeFormula <- function(
  wb,
  sheet,
  x,
  startCol = 1,
  startRow = 1,
  array = FALSE,
  xy = NULL
) {
  
  if (!is.character(x)) {
    stop("x must be a character vector.")
  }

  dfx <- data.frame("X" = x, stringsAsFactors = FALSE)
  class(dfx$X) <- c("character", ifelse(array, "array_formula", "formula"))

  if (any(grepl("^(=|)HYPERLINK\\(", x, ignore.case = TRUE))) {
    class(dfx$X) <- c("character", "formula", "hyperlink")
  }

  writeData(
    wb       = wb,
    sheet    = sheet,
    x        = dfx,
    startCol = startCol,
    startRow = startRow,
    array    = array,
    xy       = xy,
    colNames = FALSE,
    rowNames = FALSE
  )

  invisible(0)
}

#' `as.character.formula()`
#'
#' This function exists to prevent conflicts with `as.character.formula` methods
#' from other packages
#' 
#' @inheritParams base::as.character
#' @param ... Not implemented
#' @returns `base::as.character.default(x)`
#' @export
as.character.formula <- function(x, ...) {
  base::as.character.default(x)
}
