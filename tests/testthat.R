library(testthat)
library(openxlsx)

if (requireNamespace("xml2")) {
  test_check("openxlsx", reporter = MultiReporter$new(reporters = list(JunitReporter$new(file = "test-results.xml"), CheckReporter$new())))
} else {
  test_check("openxlsx")
}
