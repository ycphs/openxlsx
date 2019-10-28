openxlsx
========


[![Build Status](https://travis-ci.com/ycphs/openxlsx.svg?branch=master)](https://travis-ci.com/ycphs/openxlsx)
[![AppVeyor build status](https://ci.appveyor.com/api/projects/status/github/ycphs/openxlsx?branch=master&svg=true)](https://ci.appveyor.com/project/ycphs/openxlsx)
[![Coverage Status](https://codecov.io/github/ycphs/openxlsx/coverage.svg?branch=master)](https://codecov.io/github/ycphs/openxlsx?branch=master)
[![CRAN_Status_Badge](https://www.r-pkg.org/badges/version/openxlsx)](https://cran.r-project.org/package=openxlsx)
[![CRAN RStudio mirror downloads](https://cranlogs.r-pkg.org/badges/openxlsx)](https://cran.r-project.org/package=openxlsx)

This [R](https://www.R-project.org/) package simplifies the
creation of `.xlsx` files by providing 
a high level interface to writing, styling and editing
worksheets. Through the use of
[`Rcpp`](https://CRAN.R-project.org/package=Rcpp), 
read/write times are comparable to the
[`xlsx`](https://CRAN.R-project.org/package=xlsx)
and
[`XLConnect`](https://CRAN.R-project.org/package=XLConnect)
packages with the added benefit of removing the dependency on
Java. 

## Installation

### Stable version
Current stable version is available on
[CRAN](https://CRAN.R-project.org/) via
```R
install.packages("openxlsx", dependencies = TRUE)
```

### Development version
```R
install.packages(c("Rcpp", "devtools"), dependencies = TRUE)
require(devtools)
install_github("ycphs/openxlsx")
```

## Bug/feature request
Please let me know which version of openxlsx you are using when posting bug reports.
```R
packageVersion("openxlsx")
```
Thanks, [here](https://github.com/ycphs/openxlsx/issues). 

## News
[Here](https://raw.githubusercontent.com/ycphs/openxlsx/master/NEWS). 

## Authors and Contributors
A list is automagically maintained
[here](https://github.com/ycphs/openxlsx/graphs/contributors). 
