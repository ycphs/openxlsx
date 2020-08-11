[openxlsx](https://ycphs.github.io/openxlsx/)
========


[![Build Status](https://travis-ci.com/ycphs/openxlsx.svg?branch=master)](https://travis-ci.com/ycphs/openxlsx)
[![AppVeyor build status](https://ci.appveyor.com/api/projects/status/github/ycphs/openxlsx?branch=master&svg=true)](https://ci.appveyor.com/project/ycphs/openxlsx)
[![codecov](https://codecov.io/gh/ycphs/openxlsx/branch/master/graph/badge.svg)](https://codecov.io/gh/ycphs/openxlsx)
[![CRAN_Status_Badge](https://www.r-pkg.org/badges/version/openxlsx)](https://cran.r-project.org/package=openxlsx)
[![CRAN RStudio mirror downloads](https://cranlogs.r-pkg.org/badges/openxlsx)](https://cran.r-project.org/package=openxlsx)
![R-CMD-check](https://github.com/ycphs/openxlsx/workflows/R-CMD-check/badge.svg?branch=master)


 
 
This [R](https://www.R-project.org/) package simplifies the creation of `.xlsx` files by providing 
a high level interface to writing, styling and editing worksheets. Through the use of [`Rcpp`](https://CRAN.R-project.org/package=Rcpp), read/write times are comparable to the [`xlsx`](https://CRAN.R-project.org/package=xlsx) and
[`XLConnect`](https://CRAN.R-project.org/package=XLConnect) packages with the added benefit of removing the dependency on
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



## News
[Here](https://raw.githubusercontent.com/ycphs/openxlsx/master/NEWS.md). 


## Authors and Contributors for the current release
[&#x0040;awalker89](https://github.com/awalker89), [&#x0040;aavanesy](https://github.com/aavanesy), [&#x0040;ale275](https://github.com/ale275), [&#x0040;alexb523](https://github.com/alexb523), [&#x0040;david-f1976](https://github.com/david-f1976), [&#x0040;davidgohel](https://github.com/davidgohel), [&#x0040;dovrosenberg](https://github.com/dovrosenberg), [&#x0040;JoshuaSturm](https://github.com/JoshuaSturm), [&#x0040;SHAESEN2](https://github.com/SHAESEN2), [&#x0040;soliac](https://github.com/soliac), [&#x0040;theclue](https://github.com/theclue), and [&#x0040;ycphs](https://github.com/ycphs)
