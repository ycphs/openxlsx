% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/wrappers.R
\name{setRowHeights}
\alias{setRowHeights}
\title{Set worksheet row heights}
\usage{
setRowHeights(
  wb,
  sheet,
  rows,
  heights,
  fontsize = NULL,
  factor = 1,
  base_height = 15,
  extra_height = 12,
  wrap = TRUE
)
}
\arguments{
\item{wb}{workbook object}

\item{sheet}{name or index of a worksheet}

\item{rows}{indices of rows to set height}

\item{heights}{heights to set rows to specified in Excel column height units}

\item{fontsize}{font size, optional (get base font size by default)}

\item{factor}{factor to manually adjust font width, e.g., for bold fonts,
optional}

\item{base_height}{basic row height, optional}

\item{extra_height}{additional row height per new line of text, optional}

\item{wrap}{wrap text of entries which exceed the column width, optional}
}
\description{
Set worksheet row heights
}
\examples{
## Create a new workbook
wb <- createWorkbook()

## Add a worksheet
addWorksheet(wb, "Sheet")
sheet <- 1

## Write dummy data
writeData(wb, sheet, "fixed w/fixed h", startCol = 1, startRow = 1)
writeData(wb, sheet, "fixed w/auto h ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC",
 startCol = 2, startRow = 2)
writeData(wb, sheet, "variable w/fixed h", startCol = 3, startRow = 3)

## Set column widths and row heights
setColWidths(wb, sheet, cols = c(1, 2, 3, 4), widths = c(10, 20, "auto", 20))
setRowHeights(wb, sheet, rows = c(1, 2, 8, 4, 6), heights = c(30, "auto", 15, 15, 30))

## Overwrite row 1 height
setRowHeights(wb, sheet, rows = 1, heights = 40)

## Save workbook
\dontrun{
saveWorkbook(wb, "setRowHeightsExample.xlsx", overwrite = TRUE)
}
}
\seealso{
\code{\link[=removeRowHeights]{removeRowHeights()}}
}
\author{
Alexander Walker
}
