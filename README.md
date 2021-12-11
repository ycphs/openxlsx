[openxlsx](https://ycphs.github.io/openxlsx/) <img src="img/badge.png" align="right" height="139" />
========


[![codecov](https://codecov.io/gh/ycphs/openxlsx/branch/master/graph/badge.svg)](https://app.codecov.io/gh/ycphs/openxlsx)
[![CRAN_Status_Badge](https://www.r-pkg.org/badges/version/openxlsx)](https://cran.r-project.org/package=openxlsx)
[![CRAN RStudio mirror downloads](https://cranlogs.r-pkg.org/badges/openxlsx)](https://cran.r-project.org/package=openxlsx)
![R-CMD-check](https://github.com/ycphs/openxlsx/workflows/R-CMD-check/badge.svg?branch=master)


 
 
This [R](https://www.R-project.org/) package simplifies the creation of `.xlsx` files by providing 
a high level interface to writing, styling and editing worksheets. Through the use of [`Rcpp`](https://CRAN.R-project.org/package=Rcpp), read/write times are comparable to the [`xlsx`](https://CRAN.R-project.org/package=xlsx) and
[`XLConnect`](https://CRAN.R-project.org/package=XLConnect) packages with the added benefit of removing the dependency on
Java. 

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

## Bug/feature request
Please let me know which version of openxlsx you are using when posting bug reports.
```R
packageVersion("openxlsx")
```

## News
[Here](https://raw.githubusercontent.com/ycphs/openxlsx/master/NEWS.md). 

