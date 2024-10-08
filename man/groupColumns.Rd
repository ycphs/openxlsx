% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/wrappers.R
\name{groupColumns}
\alias{groupColumns}
\title{Group columns}
\usage{
groupColumns(wb, sheet, cols, hidden = FALSE, level = -1)
}
\arguments{
\item{wb}{A workbook object.}

\item{sheet}{A name or index of a worksheet.}

\item{cols}{Indices of cols to group. Can be either a vector of indices to
group at the same level or a (named) list of numeric vectors of
indices to create multiple groupings at once. The names of the
entries determine the grouping level. If no names are given,
the \code{level} parameter is used as default.}

\item{hidden}{Logical vector. If TRUE the grouped columns are hidden. Defaults to FALSE.}

\item{level}{Grouping level (higher value indicates multiple nestings) for the
group. A vector to assign different grouping levels to the indices.
A value of -1 indicates that the grouping level should be derived
from the existing grouping (one level added)}
}
\description{
Group a selection of columns
}
\details{
Group columns together, with the option to hide them.

NOTE: \code{\link[=setColWidths]{setColWidths()}} has a conflicting \code{hidden} parameter; changing one will update the other.
}
\seealso{
\code{\link[=ungroupColumns]{ungroupColumns()}} to ungroup columns. \code{\link[=groupRows]{groupRows()}} for grouping rows.
}
\author{
Joshua Sturm, Reinhold Kainhofer
}
