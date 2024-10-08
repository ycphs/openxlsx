% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/openxlsx.R
\docType{package}
\name{openxlsx}
\alias{openxlsx-package}
\alias{openxlsx}
\title{xlsx reading, writing and editing.}
\description{
openxlsx simplifies the the process of writing and styling Excel xlsx files from R
and removes the dependency on Java.
}
\details{
The openxlsx package uses global options, most to simplify formatting.  These
are stored in the \code{op.openxlsx} object.

\describe{
\item{openxlsx.bandedCols}{FALSE}
\item{openxlsx.bandedRows}{TRUE}
\item{openxlsx.borderColour}{"black"}
\item{openxlsx.borders}{"none"}
\item{openxlsx.borderStyle}{"thin"}
\item{openxlsx.compressionLevel}{"9"}
\item{openxlsx.creator}{""}
\item{openxlsx.dateFormat}{"mm/dd/yyyy"}
\item{openxlsx.datetimeFormat}{"yyyy-mm-dd hh:mm:ss"}
\item{openxlsx.headerStyle}{NULL}
\item{openxlsx.keepNA}{FALSE}
\item{openxlsx.na.string}{NULL}
\item{openxlsx.numFmt}{NULL}
\item{openxlsx.orientation}{"portrait"}
\item{openxlsx.paperSize}{9}
\item{openxlsx.tabColour}{"TableStyleLight9"}
\item{openxlsx.tableStyle}{"TableStyleLight9"}
\item{openxlsx.withFilter}{NA Whether to write data with or without a
filter. If NA will make filters with \code{writeDataTable} and will not for
\code{writeData}}
}

See the Formatting vignette for examples.

Additional options
}
\seealso{
\itemize{
\item{\code{vignette("Introduction", package = "openxlsx")}}
\item{\code{vignette("formatting", package = "openxlsx")}}
\item{\code{\link[=writeData]{writeData()}}}
\item{\code{\link[=writeDataTable]{writeDataTable()}}}
\item{\code{\link[=write.xlsx]{write.xlsx()}}}
\item{\code{\link[=read.xlsx]{read.xlsx()}}}
\item{\code{\link[=op.openxlsx]{op.openxlsx()}}}
}
for examples
}
\author{
\strong{Maintainer}: Jan Marvin Garbuszus \email{jan.garbuszus@ruhr-uni-bochum.de} [contributor]

Authors:
\itemize{
  \item Philipp Schauberger \email{philipp@schauberger.co.at}
  \item Alexander Walker \email{Alexander.Walker1989@gmail.com}
}

Other contributors:
\itemize{
  \item Luca Braglia [contributor]
  \item Joshua Sturm [contributor]
  \item Jordan Mark Barbone \email{jmbarbone@gmail.com} (\href{https://orcid.org/0000-0001-9788-3628}{ORCID}) [contributor]
  \item David Zimmermann \email{david_j_zimmermann@hotmail.com} [contributor]
  \item Reinhold Kainhofer \email{reinhold@kainhofer.com} [contributor]
}

}
