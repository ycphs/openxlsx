[openxlsx](https://ycphs.github.io/openxlsx/) <img src="man/figures/logo.png" align="right" height="139" />
========


[![codecov](https://codecov.io/gh/ycphs/openxlsx/branch/master/graph/badge.svg)](https://app.codecov.io/gh/ycphs/openxlsx)
[![CRAN_Status_Badge](https://www.r-pkg.org/badges/version/openxlsx)](https://cran.r-project.org/package=openxlsx)
[![CRAN RStudio mirror downloads](https://cranlogs.r-pkg.org/badges/openxlsx)](https://cran.r-project.org/package=openxlsx)
[![R-CMD-check](https://github.com/ycphs/openxlsx/actions/workflows/R-CMD-check.yaml/badge.svg?branch=master)](https://github.com/ycphs/openxlsx/actions/workflows/R-CMD-check.yaml)

This [R](https://www.R-project.org/) package simplifies the creation of `.xlsx`
files by providing a high level interface to writing, styling and editing
worksheets. Through the use of [`Rcpp`](https://CRAN.R-project.org/package=Rcpp),
read/write times are comparable to the [`xlsx`](https://CRAN.R-project.org/package=xlsx)
and [`XLConnect`](https://CRAN.R-project.org/package=XLConnect) packages with
the added benefit of removing the dependency on Java.

**Note:** `openxlsx` is no longer under active development. The package is
maintained, and CRAN warnings will be fixed, but non-critical issues will not be
addressed unless accompanied by a pull request. Packages that depend on
`openxlsx` do not need to take any action, but for new developments, users are
encouraged to use alternatives like `readxl`, `writexl`, or `openxlsx2`. The
first two packages provide support for reading and writing `.xlsx` files. The
latter package is a modern reinterpretation of `openxlsx` and provides similar
functions to modify worksheets. However, it is not a drop-in replacement, so you
may want to consult resources like the
[update vignette](https://janmarvin.github.io/openxlsx2/articles/Update-from-openxlsx.html).


## Installation

### Stable version

Current stable version is available on [CRAN](https://CRAN.R-project.org/) via

```R
install.packages("openxlsx", dependencies = TRUE)
```

### Development version
```R
install.packages(c("Rcpp", "remotes"), dependencies = TRUE)
remotes::install_github("ycphs/openxlsx")
```

## Example

Explore the package with a simple example:

```R
library(openxlsx)

# Create a new workbook and add a sheet
wb <- createWorkbook()
addWorksheet(wb, "Sheet 1")

# Write data to the sheet
writeData(wb, "Sheet 1", mtcars)

# Save the workbook
saveWorkbook(wb, "my_mtcars.xlsx", overwrite = TRUE)
```

## Bug/feature request
Please let us know which version of `openxlsx` you are using when posting bug reports.
```R
packageVersion("openxlsx")
```

## News
You can find the NEWS file [here](https://raw.githubusercontent.com/ycphs/openxlsx/master/NEWS.md).
