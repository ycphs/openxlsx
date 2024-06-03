




#' @name conditionalFormatting
#' @aliases databar
#' @title Add conditional formatting to cells
#' @description Add conditional formatting to cells
#' @author Alexander Walker, Philipp Schauberger
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols Columns to apply conditional formatting to
#' @param rows Rows to apply conditional formatting to
#' @param rule The condition under which to apply the formatting. See examples.
#' @param style A style to apply to those cells that satisfy the rule. Default is createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
#' @param type Either 'expression', 'colourScale', 'databar', 'duplicates', 'beginsWith', 
#' 'endsWith', 'topN', 'bottomN', 'blanks', 'notBlanks', 'contains' or 'notContains' (case insensitive).
#' @param ... See below
#' @details See Examples.
#'
#' If type == "expression"
#' \itemize{
#'   \item{style is a Style object. See [createStyle()]}
#'   \item{rule is an expression. Valid operators are "<", "<=", ">", ">=", "==", "!=".}
#' }
#'
#' If type == "colourScale"
#' \itemize{
#'   \item{style is a vector of colours with length 2 or 3}
#'   \item{rule can be NULL or a vector of colours of equal length to styles}
#' }
#'
#' If type == "databar"
#' \itemize{
#'   \item{style is a vector of colours with length 2 or 3}
#'   \item{rule is a numeric vector specifying the range of the databar colours. Must be equal length to style}
#'   \item{...
#'   \itemize{
#'     \item{**showvalue** If FALSE the cell value is hidden. Default TRUE.}
#'     \item{**gradient** If FALSE colour gradient is removed. Default TRUE.}
#'     \item{**border** If FALSE the border around the database is hidden. Default TRUE.}
#'      }
#'    }
#' }
#'
#' If type == "duplicates"
#' \itemize{
#'   \item{style is a Style object. See [createStyle()]}
#'   \item{rule is ignored.}
#' }
#'
#' If type == "contains"
#' \itemize{
#'   \item{style is a Style object. See [createStyle()]}
#'   \item{rule is the text to look for within cells}
#' }
#'
#' If type == "between"
#' \itemize{
#'   \item{style is a Style object. See [createStyle()]}
#'   \item{rule is a numeric vector of length 2 specifying lower and upper bound (Inclusive)}
#' }
#' If type == "blanks"
#' \itemize{
#'   \item{style is a Style object. See [createStyle()]}
#'   \item{rule is ignored.}
#' }
#'
#' If type == "notBlanks"
#' \itemize{
#'   \item{style is a Style object. See [createStyle()]}
#'   \item{rule is ignored.}
#' }

#'
#' If type == "topN"
#' \itemize{
#'   \item{style is a Style object. See [createStyle()]}
#'   \item{rule is ignored}
#'   \item{...
#'   \itemize{
#'     \item{**rank** numeric vector of length 1 indicating number of highest values.}
#'     \item{**percent** TRUE if you want top N percentage.}
#'      }
#'    }
#' }
#' 
#' If type == "bottomN"
#' \itemize{
#'   \item{style is a Style object. See [createStyle()]}
#'   \item{rule is ignored}
#'   \item{...
#'   \itemize{
#'     \item{**rank** numeric vector of length 1 indicating number of lowest values.}
#'     \item{**percent** TRUE if you want bottom N percentage.}
#'      }
#'    }
#' }
#' 
#' @seealso [createStyle()]
#' @export
#' @examples
#' wb <- createWorkbook()
#' addWorksheet(wb, "cellIs")
#' addWorksheet(wb, "Moving Row")
#' addWorksheet(wb, "Moving Col")
#' addWorksheet(wb, "Dependent on")
#' addWorksheet(wb, "Duplicates")
#' addWorksheet(wb, "containsText")
#' addWorksheet(wb, "notcontainsText")
#' addWorksheet(wb, "beginsWith")
#' addWorksheet(wb, "endsWith")
#' addWorksheet(wb, "colourScale", zoom = 30)
#' addWorksheet(wb, "databar")
#' addWorksheet(wb, "between")
#' addWorksheet(wb, "topN")
#' addWorksheet(wb, "bottomN")
#' addWorksheet(wb, "containsBlanks")
#' addWorksheet(wb, "notContainsBlanks")
#' addWorksheet(wb, "logical operators")
#'
#' negStyle <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
#' posStyle <- createStyle(fontColour = "#006100", bgFill = "#C6EFCE")
#'
#' ## rule applies to all each cell in range
#' writeData(wb, "cellIs", -5:5)
#' writeData(wb, "cellIs", LETTERS[1:11], startCol = 2)
#' conditionalFormatting(wb, "cellIs",
#'   cols = 1,
#'   rows = 1:11, rule = "!=0", style = negStyle
#' )
#' conditionalFormatting(wb, "cellIs",
#'   cols = 1,
#'   rows = 1:11, rule = "==0", style = posStyle
#' )
#'
#' ## highlight row dependent on first cell in row
#' writeData(wb, "Moving Row", -5:5)
#' writeData(wb, "Moving Row", LETTERS[1:11], startCol = 2)
#' conditionalFormatting(wb, "Moving Row",
#'   cols = 1:2,
#'   rows = 1:11, rule = "$A1<0", style = negStyle
#' )
#' conditionalFormatting(wb, "Moving Row",
#'   cols = 1:2,
#'   rows = 1:11, rule = "$A1>0", style = posStyle
#' )
#'
#' ## highlight column dependent on first cell in column
#' writeData(wb, "Moving Col", -5:5)
#' writeData(wb, "Moving Col", LETTERS[1:11], startCol = 2)
#' conditionalFormatting(wb, "Moving Col",
#'   cols = 1:2,
#'   rows = 1:11, rule = "A$1<0", style = negStyle
#' )
#' conditionalFormatting(wb, "Moving Col",
#'   cols = 1:2,
#'   rows = 1:11, rule = "A$1>0", style = posStyle
#' )
#'
#' ## highlight entire range cols X rows dependent only on cell A1
#' writeData(wb, "Dependent on", -5:5)
#' writeData(wb, "Dependent on", LETTERS[1:11], startCol = 2)
#' conditionalFormatting(wb, "Dependent on",
#'   cols = 1:2,
#'   rows = 1:11, rule = "$A$1<0", style = negStyle
#' )
#' conditionalFormatting(wb, "Dependent on",
#'   cols = 1:2,
#'   rows = 1:11, rule = "$A$1>0", style = posStyle
#' )
#'
#' ## highlight cells in column 1 based on value in column 2
#' writeData(wb, "Dependent on", data.frame(x = 1:10, y = runif(10)), startRow = 15)
#' conditionalFormatting(wb, "Dependent on",
#'   cols = 1,
#'   rows = 16:25, rule = "B16<0.5", style = negStyle
#' )
#' conditionalFormatting(wb, "Dependent on",
#'   cols = 1,
#'   rows = 16:25, rule = "B16>=0.5", style = posStyle
#' )
#'
#'
#' ## highlight duplicates using default style
#' writeData(wb, "Duplicates", sample(LETTERS[1:15], size = 10, replace = TRUE))
#' conditionalFormatting(wb, "Duplicates", cols = 1, rows = 1:10, type = "duplicates")
#'
#' ## cells containing text
#' fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
#' writeData(wb, "containsText", sapply(1:10, fn))
#' conditionalFormatting(wb, "containsText", cols = 1, rows = 1:10, type = "contains", rule = "A")
#' 
#' ## cells not containing text
#' fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
#' writeData(wb, "containsText", sapply(1:10, fn))
#' conditionalFormatting(wb, "notcontainsText", cols = 1, 
#'                      rows = 1:10, type = "notcontains", rule = "A")
#' 
#' 
#' ## cells begins with text
#' fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
#' writeData(wb, "beginsWith", sapply(1:100, fn))
#' conditionalFormatting(wb, "beginsWith", cols = 1, rows = 1:100, type = "beginsWith", rule = "A")
#' 
#' 
#' ## cells ends with text
#' fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
#' writeData(wb, "endsWith", sapply(1:100, fn))
#' conditionalFormatting(wb, "endsWith", cols = 1, rows = 1:100, type = "endsWith", rule = "A")
#'
#' ## colourscale colours cells based on cell value
#' df <- read.xlsx(system.file("extdata", "readTest.xlsx", package = "openxlsx"), sheet = 4)
#' writeData(wb, "colourScale", df, colNames = FALSE) ## write data.frame
#'
#' ## rule is a vector or colours of length 2 or 3 (any hex colour or any of colours())
#' ## If rule is NULL, min and max of cells is used. Rule must be the same length as style or NULL.
#' conditionalFormatting(wb, "colourScale",
#'   cols = 1:ncol(df), rows = 1:nrow(df),
#'   style = c("black", "white"),
#'   rule = c(0, 255),
#'   type = "colourScale"
#' )
#'
#' setColWidths(wb, "colourScale", cols = 1:ncol(df), widths = 1.07)
#' setRowHeights(wb, "colourScale", rows = 1:nrow(df), heights = 7.5)
#'
#' ## Databars
#' writeData(wb, "databar", -5:5)
#' conditionalFormatting(wb, "databar", cols = 1, rows = 1:11, type = "databar") ## Default colours
#'
#' ## Between
#' # Highlight cells in interval [-2, 2]
#' writeData(wb, "between", -5:5)
#' conditionalFormatting(wb, "between", cols = 1, rows = 1:11, type = "between", rule = c(-2, 2))
#'
#' ## Top N 
#' writeData(wb, "topN", data.frame(x = 1:10, y = rnorm(10)))
#' # Highlight top 5 values in column x
#' conditionalFormatting(wb, "topN", cols = 1, rows = 2:11, 
#'  style = posStyle, type = "topN", rank = 5)#'
#' # Highlight top 20 percentage in column y
#' conditionalFormatting(wb, "topN", cols = 2, rows = 2:11, 
#'  style = posStyle, type = "topN", rank = 20, percent = TRUE)
#'
#'## Bottom N 
#' writeData(wb, "bottomN", data.frame(x = 1:10, y = rnorm(10)))
#' # Highlight bottom 5 values in column x
#' conditionalFormatting(wb, "bottomN", cols = 1, rows = 2:11, 
#'  style = negStyle, type = "topN", rank = 5)
#' # Highlight bottom 20 percentage in column y
#' conditionalFormatting(wb, "bottomN", cols = 2, rows = 2:11, 
#'  style = negStyle, type = "topN", rank = 20, percent = TRUE)
#'  
#' ## cells containing blanks
#' sample_data <- sample(c("X", NA_character_), 10, replace = TRUE)
#' writeData(wb, "containsBlanks", sample_data)
#' conditionalFormatting(wb, "containsBlanks", cols = 1, rows = 1:10, 
#' type = "blanks", style = negStyle)
#' 
#' ## cells not containing blanks
#' sample_data <- sample(c("X", NA_character_), 10, replace = TRUE)
#' writeData(wb, "notContainsBlanks", sample_data)
#' conditionalFormatting(wb, "notContainsBlanks", cols = 1, rows = 1:10, 
#' type = "notBlanks", style = posStyle)
#'
#' ## Logical Operators
#' # You can use Excels logical Operators
#' writeData(wb, "logical operators", 1:10)
#' conditionalFormatting(wb, "logical operators",
#'   cols = 1, rows = 1:10,
#'   rule = "OR($A1=1,$A1=3,$A1=5,$A1=7)"
#' )
#' \dontrun{
#' saveWorkbook(wb, "conditionalFormattingExample.xlsx", TRUE)
#' }
#'
#'
#' #########################################################################
#' ## Databar Example
#'
#' wb <- createWorkbook()
#' addWorksheet(wb, "databar")
#'
#' ## Databars
#' writeData(wb, "databar", -5:5, startCol = 1)
#' conditionalFormatting(wb, "databar", cols = 1, rows = 1:11, type = "databar") ## Defaults
#'
#' writeData(wb, "databar", -5:5, startCol = 3)
#' conditionalFormatting(wb, "databar", cols = 3, rows = 1:11, type = "databar", border = FALSE)
#'
#' writeData(wb, "databar", -5:5, startCol = 5)
#' conditionalFormatting(wb, "databar",
#'   cols = 5, rows = 1:11,
#'   type = "databar", style = c("#a6a6a6"), showValue = FALSE
#' )
#'
#' writeData(wb, "databar", -5:5, startCol = 7)
#' conditionalFormatting(wb, "databar",
#'   cols = 7, rows = 1:11,
#'   type = "databar", style = c("#a6a6a6"), showValue = FALSE, gradient = FALSE
#' )
#'
#' writeData(wb, "databar", -5:5, startCol = 9)
#' conditionalFormatting(wb, "databar",
#'   cols = 9, rows = 1:11,
#'   type = "databar", style = c("#a6a6a6", "#a6a6a6"), showValue = FALSE, gradient = FALSE
#' )
#' \dontrun{
#' saveWorkbook(wb, file = "databarExample.xlsx", overwrite = TRUE)
#' }
#'
conditionalFormatting <-
  function(wb,
           sheet,
           cols,
           rows,
           rule = NULL,
           style = NULL,
           type = "expression",
           ...) {
    op <- get_set_options()
    on.exit(options(op), add = TRUE)

    type <- tolower(type)
    params <- list(...)

    if (type %in% c("colorscale", "colourscale")) {
      type <- "colorScale"
    } else if (type == "databar") {
      type <- "dataBar"
    } else if (type == "duplicates") {
      type <- "duplicatedValues"
    } else if (type == "contains") {
      type <- "containsText"
    } else if (type == "notcontains") {
      type <- "notContainsText"
    } else if (type == "beginswith") {
      type <- "beginsWith"
    } else if (type == "endswith") {
      type <- "endsWith"
    } else if (type == "between") {
      type <- "between"
    } else if (type == "topn") {
      type <- "topN"
    } else if (type == "bottomn") {
      type <- "bottomN"
    } else if (type == "blanks") {
      type <- "containsBlanks"
    } else if (type == "notblanks") {
      type <- "notContainsBlanks"
    }else if (type != "expression") {
      stop(
        "Invalid type argument.  Type must be one of 'expression', 'colourScale', 'databar', 'duplicates', 'beginsWith', 'endsWith', 'blanks', 'notBlanks', 'contains' or 'notContains'"
      )
    }

    ## rows and cols
    if (!is.numeric(cols)) {
      cols <- convertFromExcelRef(cols)
    }
    rows <- as.integer(rows)


    ## check valid rule
    values <- NULL
    dxfId <- NULL

    if (type == "colorScale") {
      # type == "colourScale"
      # - style is a vector of colours with length 2 or 3
      # - rule specifies the quantiles (numeric vector of length 2 or 3), if NULL min and max are used

      if (is.null(style)) {
        stop("If type == 'colourScale', style must be a vector of colours of length 2 or 3.")
      }

      if (!inherits(style, "character")) {
        stop("If type == 'colourScale', style must be a vector of colours of length 2 or 3.")
      }

      if (!length(style) %in% 2:3) {
        stop("If type == 'colourScale', style must be a vector of length 2 or 3.")
      }

      if (!is.null(rule)) {
        if (length(rule) != length(style)) {
          stop("If type == 'colourScale', rule and style must have equal lengths.")
        }
      }

      style <-
        validateColour(style, errorMsg = "Invalid colour specified in style.")

      values <- rule
      rule <- style
    } else if (type == "dataBar") {
      # type == "databar"
      # - style is a vector of colours of length 2 or 3
      # - rule specifies the quantiles (numeric vector of length 2 or 3), if NULL min and max are used

      if (is.null(style)) {
        style <- "#638EC6"
      }

      if (!inherits(style, "character")) {
        stop("If type == 'dataBar', style must be a vector of colours of length 1 or 2.")
      }

      if (!length(style) %in% 1:2) {
        stop("If type == 'dataBar', style must be a vector of length 1 or 2.")
      }

      if (!is.null(rule)) {
        if (length(rule) != length(style)) {
          stop("If type == 'dataBar', rule and style must have equal lengths.")
        }
      }


      ## Additional parameters passed by ...
      if ("showValue" %in% names(params)) {
        params$showValue <- as.integer(params$showValue)
        if (is.na(params$showValue)) {
          stop("showValue must be 0/1 or TRUE/FALSE")
        }
      }

      if ("gradient" %in% names(params)) {
        params$gradient <- as.integer(params$gradient)
        if (is.na(params$gradient)) {
          stop("gradient must be 0/1 or TRUE/FALSE")
        }
      }

      if ("border" %in% names(params)) {
        params$border <- as.integer(params$border)
        if (is.na(params$border)) {
          stop("border must be 0/1 or TRUE/FALSE")
        }
      }

      style <-
        validateColour(style, errorMsg = "Invalid colour specified in style.")

      values <- rule
      rule <- style
    } else if (type == "expression") {
      # type == "expression"
      # - style = createStyle()
      # - rule is an expression to evaluate

      # rule <- gsub(" ", "", rule)
      rule <- replaceIllegalCharacters(rule)
      rule <- gsub("!=", "&lt;&gt;", rule)
      rule <- gsub("==", "=", rule)

      if (!grepl("[A-Z]", substr(rule, 1, 2))) {
        ## formula looks like "operatorX" , attach top left cell to rule
        rule <-
          paste0(getCellRefs(data.frame(
            "x" = min(rows), "y" = min(cols)
          )), rule)
      } ## else, there is a letter in the formula and apply as is

      if (is.null(style)) {
        style <-
          createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
      }

      if (!"Style" %in% class(style)) {
        stop("If type == 'expression', style must be a Style object.")
      }

      invisible(dxfId <- wb$addDXFS(style))
    } else if (type == "duplicatedValues") {
      # type == "duplicatedValues"
      # - style is a Style object
      # - rule is ignored

      if (is.null(style)) {
        style <-
          createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
      }

      if (!"Style" %in% class(style)) {
        stop("If type == 'duplicates', style must be a Style object.")
      }

      invisible(dxfId <- wb$addDXFS(style))
      rule <- style
    } else if (type == "containsText") {
      # type == "contains"
      # - style is Style object
      # - rule is text to look for
      
      if (is.null(style)) {
        style <-
          createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
      }
      
      
      if (!"character" %in% class(rule)) {
        stop("If type == 'contains', rule must be a character vector of length 1.")
      }
      
      if (!"Style" %in% class(style)) {
        stop("If type == 'contains', style must be a Style object.")
      }
      
      invisible(dxfId <- wb$addDXFS(style))
      values <- rule
      rule <- style
    } else if (type == "notContainsText") {
      # type == "contains"
      # - style is Style object
      # - rule is text to look for
      
      if (is.null(style)) {
        style <-
          createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
      }
      
      
      if (!"character" %in% class(rule)) {
        stop("If type == 'notContains', rule must be a character vector of length 1.")
      }
      
      if (!"Style" %in% class(style)) {
        stop("If type == 'notContains', style must be a Style object.")
      }
      
      invisible(dxfId <- wb$addDXFS(style))
      values <- rule
      rule <- style
    } else if (type == "beginsWith") {
      # type == "contains"
      # - style is Style object
      # - rule is text to look for
      
      if (is.null(style)) {
        style <-
          createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
      }
      
      
      if (!"character" %in% class(rule)) {
        stop("If type == 'beginsWith', rule must be a character vector of length 1.")
      }
      
      if (!"Style" %in% class(style)) {
        stop("If type == 'beginsWith', style must be a Style object.")
      }
      
      invisible(dxfId <- wb$addDXFS(style))
      values <- rule
      rule <- style
    } else if (type == "endsWith") {
      # type == "contains"
      # - style is Style object
      # - rule is text to look for
      
      if (is.null(style)) {
        style <-
          createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
      }
      
      
      if (!"character" %in% class(rule)) {
        stop("If type == 'endsWith', rule must be a character vector of length 1.")
      }
      
      if (!"Style" %in% class(style)) {
        stop("If type == 'endsWith', style must be a Style object.")
      }
      
      invisible(dxfId <- wb$addDXFS(style))
      values <- rule
      rule <- style
    } else if (type == "between") {
      rule <- range(rule)

      if (is.null(style)) {
        style <-
          createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
      }

      if (!"Style" %in% class(style)) {
        stop("If type == 'between', style must be a Style object.")
      }

      invisible(dxfId <- wb$addDXFS(style))
    } else if (type == "topN") {
      # type == "topN"
      # - rule is ignored
      # - 'rank' and 'percent' are named params 
      
      if (is.null(style)) {
        style <-
          createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
      } 
      
      if (!"Style" %in% class(style)) {
        stop("If type == 'topN', style must be a Style object.")
      } 
      
      invisible(dxfId <- wb$addDXFS(style)) 
      
      ## Additional parameters passed by ...
      if ("percent" %in% names(params)) {
        params$percent <- as.integer(params$percent)
        if (is.na(params$percent)) {
          stop("percent must be 0/1 or TRUE/FALSE")
        }
      } 
      
      if ("rank" %in% names(params)) {
        params$rank <- as.integer(params$rank)
        if (is.na(params$rank)) {
          stop("rank must be a number")
        }
      } 
      
      invisible(dxfId <- wb$addDXFS(style))
      values <- params
      rule <- style
    } else if (type == "bottomN") {
      # type == "bottomN"
      # - rule is ignored
      # - 'rank' and 'percent' are named params 
      
      if (is.null(style)) {
        style <-
          createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
      } 
      
      if (!"Style" %in% class(style)) {
        stop("If type == 'bottomN', style must be a Style object.")
      }
      
      invisible(dxfId <- wb$addDXFS(style)) 
      
      ## Additional parameters passed by ...
      if ("percent" %in% names(params)) {
        params$percent <- as.integer(params$percent)
        if (is.na(params$percent)) {
          stop("percent must be 0/1 or TRUE/FALSE")
        }
      } 
      
      if ("rank" %in% names(params)) {
        params$rank <- as.integer(params$rank)
        if (is.na(params$rank)) {
          stop("rank must be a number")
        }
      } 
      
      invisible(dxfId <- wb$addDXFS(style))
      values <- params
      rule <- style
    }  else if (type == "containsBlanks") {
      # rule is ignored
      
      if (is.null(style)) {
        style <-
          createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
      }
      
      if (!"Style" %in% class(style)) {
        stop("If type == 'blanks', style must be a Style object.")
      }
      
      invisible(dxfId <- wb$addDXFS(style))
    } else if (type == "notContainsBlanks") {
      # rule is ignored
      
      if (is.null(style)) {
        style <-
          createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
      }
      
      if (!"Style" %in% class(style)) {
        stop("If type == 'notBlanks', style must be a Style object.")
      }
      
      invisible(dxfId <- wb$addDXFS(style))
    }



    invisible(
      wb$conditionalFormatting(
        sheet,
        startRow = min(rows),
        endRow = max(rows),
        startCol = min(cols),
        endCol = max(cols),
        dxfId = dxfId,
        formula = rule,
        type = type,
        values = values,
        params = params
      )
    )

    invisible(0)
  }
