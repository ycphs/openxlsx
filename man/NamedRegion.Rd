% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/wrappers.R
\name{createNamedRegion}
\alias{createNamedRegion}
\alias{deleteNamedRegion}
\title{Create / delete a named region.}
\usage{
createNamedRegion(wb, sheet, cols, rows, name, overwrite = FALSE)

deleteNamedRegion(wb, name)
}
\arguments{
\item{wb}{A workbook object}

\item{sheet}{A name or index of a worksheet}

\item{cols}{Numeric vector specifying columns to include in region}

\item{rows}{Numeric vector specifying rows to include in region}

\item{name}{Name for region. A character vector of length 1. Note region names must be case-insensitive unique.}

\item{overwrite}{Boolean. Overwrite if exists ? Default to FALSE}
}
\description{
Create / delete a named region
}
\details{
Region is given by: min(cols):max(cols) X min(rows):max(rows)
}
\examples{
## create named regions
wb <- createWorkbook()
addWorksheet(wb, "Sheet 1")

## specify region
writeData(wb, sheet = 1, x = iris, startCol = 1, startRow = 1)
createNamedRegion(
  wb = wb,
  sheet = 1,
  name = "iris",
  rows = 1:(nrow(iris) + 1),
  cols = 1:ncol(iris)
)


## using writeData 'name' argument
writeData(wb, sheet = 1, x = iris, name = "iris2", startCol = 10)

out_file <- tempfile(fileext = ".xlsx")
\dontrun{
saveWorkbook(wb, out_file, overwrite = TRUE)

## see named regions
getNamedRegions(wb) ## From Workbook object
getNamedRegions(out_file) ## From xlsx file

## delete one
deleteNamedRegion(wb = wb, name = "iris2")
getNamedRegions(wb)

## read named regions
df <- read.xlsx(wb, namedRegion = "iris")
head(df)

df <- read.xlsx(out_file, namedRegion = "iris2")
head(df)
}

}
\seealso{
\code{\link[=getNamedRegions]{getNamedRegions()}}
}
\author{
Alexander Walker
}
