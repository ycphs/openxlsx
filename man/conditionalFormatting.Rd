% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/conditional_formatting.R
\name{conditionalFormatting}
\alias{conditionalFormatting}
\alias{databar}
\title{Add conditional formatting to cells}
\usage{
conditionalFormatting(
  wb,
  sheet,
  cols,
  rows,
  rule = NULL,
  style = NULL,
  type = "expression",
  ...
)
}
\arguments{
\item{wb}{A workbook object}

\item{sheet}{A name or index of a worksheet}

\item{cols}{Columns to apply conditional formatting to}

\item{rows}{Rows to apply conditional formatting to}

\item{rule}{The condition under which to apply the formatting. See examples.}

\item{style}{A style to apply to those cells that satisfy the rule. Default is createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")}

\item{type}{Either 'expression', 'colourScale', 'databar', 'duplicates', 'beginsWith',
'endsWith', 'topN', 'bottomN', 'blanks', 'notBlanks', 'contains' or 'notContains' (case insensitive).}

\item{...}{See below}
}
\description{
Add conditional formatting to cells
}
\details{
See Examples.

If type == "expression"
\itemize{
\item{style is a Style object. See \code{\link[=createStyle]{createStyle()}}}
\item{rule is an expression. Valid operators are "<", "<=", ">", ">=", "==", "!=".}
}

If type == "colourScale"
\itemize{
\item{style is a vector of colours with length 2 or 3}
\item{rule can be NULL or a vector of colours of equal length to styles}
}

If type == "databar"
\itemize{
\item{style is a vector of colours with length 2 or 3}
\item{rule is a numeric vector specifying the range of the databar colours. Must be equal length to style}
\item{...
\itemize{
\item{\strong{showvalue} If FALSE the cell value is hidden. Default TRUE.}
\item{\strong{gradient} If FALSE colour gradient is removed. Default TRUE.}
\item{\strong{border} If FALSE the border around the database is hidden. Default TRUE.}
}
}
}

If type == "duplicates"
\itemize{
\item{style is a Style object. See \code{\link[=createStyle]{createStyle()}}}
\item{rule is ignored.}
}

If type == "contains"
\itemize{
\item{style is a Style object. See \code{\link[=createStyle]{createStyle()}}}
\item{rule is the text to look for within cells}
}

If type == "between"
\itemize{
\item{style is a Style object. See \code{\link[=createStyle]{createStyle()}}}
\item{rule is a numeric vector of length 2 specifying lower and upper bound (Inclusive)}
}
If type == "blanks"
\itemize{
\item{style is a Style object. See \code{\link[=createStyle]{createStyle()}}}
\item{rule is ignored.}
}

If type == "notBlanks"
\itemize{
\item{style is a Style object. See \code{\link[=createStyle]{createStyle()}}}
\item{rule is ignored.}
}

If type == "topN"
\itemize{
\item{style is a Style object. See \code{\link[=createStyle]{createStyle()}}}
\item{rule is ignored}
\item{...
\itemize{
\item{\strong{rank} numeric vector of length 1 indicating number of highest values.}
\item{\strong{percent} TRUE if you want top N percentage.}
}
}
}

If type == "bottomN"
\itemize{
\item{style is a Style object. See \code{\link[=createStyle]{createStyle()}}}
\item{rule is ignored}
\item{...
\itemize{
\item{\strong{rank} numeric vector of length 1 indicating number of lowest values.}
\item{\strong{percent} TRUE if you want bottom N percentage.}
}
}
}
}
\examples{
wb <- createWorkbook()
addWorksheet(wb, "cellIs")
addWorksheet(wb, "Moving Row")
addWorksheet(wb, "Moving Col")
addWorksheet(wb, "Dependent on")
addWorksheet(wb, "Duplicates")
addWorksheet(wb, "containsText")
addWorksheet(wb, "notcontainsText")
addWorksheet(wb, "beginsWith")
addWorksheet(wb, "endsWith")
addWorksheet(wb, "colourScale", zoom = 30)
addWorksheet(wb, "databar")
addWorksheet(wb, "between")
addWorksheet(wb, "topN")
addWorksheet(wb, "bottomN")
addWorksheet(wb, "containsBlanks")
addWorksheet(wb, "notContainsBlanks")
addWorksheet(wb, "logical operators")

negStyle <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
posStyle <- createStyle(fontColour = "#006100", bgFill = "#C6EFCE")

## rule applies to all each cell in range
writeData(wb, "cellIs", -5:5)
writeData(wb, "cellIs", LETTERS[1:11], startCol = 2)
conditionalFormatting(wb, "cellIs",
  cols = 1,
  rows = 1:11, rule = "!=0", style = negStyle
)
conditionalFormatting(wb, "cellIs",
  cols = 1,
  rows = 1:11, rule = "==0", style = posStyle
)

## highlight row dependent on first cell in row
writeData(wb, "Moving Row", -5:5)
writeData(wb, "Moving Row", LETTERS[1:11], startCol = 2)
conditionalFormatting(wb, "Moving Row",
  cols = 1:2,
  rows = 1:11, rule = "$A1<0", style = negStyle
)
conditionalFormatting(wb, "Moving Row",
  cols = 1:2,
  rows = 1:11, rule = "$A1>0", style = posStyle
)

## highlight column dependent on first cell in column
writeData(wb, "Moving Col", -5:5)
writeData(wb, "Moving Col", LETTERS[1:11], startCol = 2)
conditionalFormatting(wb, "Moving Col",
  cols = 1:2,
  rows = 1:11, rule = "A$1<0", style = negStyle
)
conditionalFormatting(wb, "Moving Col",
  cols = 1:2,
  rows = 1:11, rule = "A$1>0", style = posStyle
)

## highlight entire range cols X rows dependent only on cell A1
writeData(wb, "Dependent on", -5:5)
writeData(wb, "Dependent on", LETTERS[1:11], startCol = 2)
conditionalFormatting(wb, "Dependent on",
  cols = 1:2,
  rows = 1:11, rule = "$A$1<0", style = negStyle
)
conditionalFormatting(wb, "Dependent on",
  cols = 1:2,
  rows = 1:11, rule = "$A$1>0", style = posStyle
)

## highlight cells in column 1 based on value in column 2
writeData(wb, "Dependent on", data.frame(x = 1:10, y = runif(10)), startRow = 15)
conditionalFormatting(wb, "Dependent on",
  cols = 1,
  rows = 16:25, rule = "B16<0.5", style = negStyle
)
conditionalFormatting(wb, "Dependent on",
  cols = 1,
  rows = 16:25, rule = "B16>=0.5", style = posStyle
)


## highlight duplicates using default style
writeData(wb, "Duplicates", sample(LETTERS[1:15], size = 10, replace = TRUE))
conditionalFormatting(wb, "Duplicates", cols = 1, rows = 1:10, type = "duplicates")

## cells containing text
fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
writeData(wb, "containsText", sapply(1:10, fn))
conditionalFormatting(wb, "containsText", cols = 1, rows = 1:10, type = "contains", rule = "A")

## cells not containing text
fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
writeData(wb, "containsText", sapply(1:10, fn))
conditionalFormatting(wb, "notcontainsText", cols = 1, 
                     rows = 1:10, type = "notcontains", rule = "A")


## cells begins with text
fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
writeData(wb, "beginsWith", sapply(1:100, fn))
conditionalFormatting(wb, "beginsWith", cols = 1, rows = 1:100, type = "beginsWith", rule = "A")


## cells ends with text
fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
writeData(wb, "endsWith", sapply(1:100, fn))
conditionalFormatting(wb, "endsWith", cols = 1, rows = 1:100, type = "endsWith", rule = "A")

## colourscale colours cells based on cell value
df <- read.xlsx(system.file("extdata", "readTest.xlsx", package = "openxlsx"), sheet = 4)
writeData(wb, "colourScale", df, colNames = FALSE) ## write data.frame

## rule is a vector or colours of length 2 or 3 (any hex colour or any of colours())
## If rule is NULL, min and max of cells is used. Rule must be the same length as style or NULL.
conditionalFormatting(wb, "colourScale",
  cols = 1:ncol(df), rows = 1:nrow(df),
  style = c("black", "white"),
  rule = c(0, 255),
  type = "colourScale"
)

setColWidths(wb, "colourScale", cols = 1:ncol(df), widths = 1.07)
setRowHeights(wb, "colourScale", rows = 1:nrow(df), heights = 7.5)

## Databars
writeData(wb, "databar", -5:5)
conditionalFormatting(wb, "databar", cols = 1, rows = 1:11, type = "databar") ## Default colours

## Between
# Highlight cells in interval [-2, 2]
writeData(wb, "between", -5:5)
conditionalFormatting(wb, "between", cols = 1, rows = 1:11, type = "between", rule = c(-2, 2))

## Top N 
writeData(wb, "topN", data.frame(x = 1:10, y = rnorm(10)))
# Highlight top 5 values in column x
conditionalFormatting(wb, "topN", cols = 1, rows = 2:11, 
 style = posStyle, type = "topN", rank = 5)#'
# Highlight top 20 percentage in column y
conditionalFormatting(wb, "topN", cols = 2, rows = 2:11, 
 style = posStyle, type = "topN", rank = 20, percent = TRUE)

## Bottom N 
writeData(wb, "bottomN", data.frame(x = 1:10, y = rnorm(10)))
# Highlight bottom 5 values in column x
conditionalFormatting(wb, "bottomN", cols = 1, rows = 2:11, 
 style = negStyle, type = "topN", rank = 5)
# Highlight bottom 20 percentage in column y
conditionalFormatting(wb, "bottomN", cols = 2, rows = 2:11, 
 style = negStyle, type = "topN", rank = 20, percent = TRUE)
 
## cells containing blanks
sample_data <- sample(c("X", NA_character_), 10, replace = TRUE)
writeData(wb, "containsBlanks", sample_data)
conditionalFormatting(wb, "containsBlanks", cols = 1, rows = 1:10, 
type = "blanks", style = negStyle)

## cells not containing blanks
sample_data <- sample(c("X", NA_character_), 10, replace = TRUE)
writeData(wb, "notContainsBlanks", sample_data)
conditionalFormatting(wb, "notContainsBlanks", cols = 1, rows = 1:10, 
type = "notBlanks", style = posStyle)

## Logical Operators
# You can use Excels logical Operators
writeData(wb, "logical operators", 1:10)
conditionalFormatting(wb, "logical operators",
  cols = 1, rows = 1:10,
  rule = "OR($A1=1,$A1=3,$A1=5,$A1=7)"
)
\dontrun{
saveWorkbook(wb, "conditionalFormattingExample.xlsx", TRUE)
}


#########################################################################
## Databar Example

wb <- createWorkbook()
addWorksheet(wb, "databar")

## Databars
writeData(wb, "databar", -5:5, startCol = 1)
conditionalFormatting(wb, "databar", cols = 1, rows = 1:11, type = "databar") ## Defaults

writeData(wb, "databar", -5:5, startCol = 3)
conditionalFormatting(wb, "databar", cols = 3, rows = 1:11, type = "databar", border = FALSE)

writeData(wb, "databar", -5:5, startCol = 5)
conditionalFormatting(wb, "databar",
  cols = 5, rows = 1:11,
  type = "databar", style = c("#a6a6a6"), showValue = FALSE
)

writeData(wb, "databar", -5:5, startCol = 7)
conditionalFormatting(wb, "databar",
  cols = 7, rows = 1:11,
  type = "databar", style = c("#a6a6a6"), showValue = FALSE, gradient = FALSE
)

writeData(wb, "databar", -5:5, startCol = 9)
conditionalFormatting(wb, "databar",
  cols = 9, rows = 1:11,
  type = "databar", style = c("#a6a6a6", "#a6a6a6"), showValue = FALSE, gradient = FALSE
)
\dontrun{
saveWorkbook(wb, file = "databarExample.xlsx", overwrite = TRUE)
}

}
\seealso{
\code{\link[=createStyle]{createStyle()}}
}
\author{
Alexander Walker, Philipp Schauberger
}
