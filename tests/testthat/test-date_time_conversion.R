
context("Date/Time Conversions")

test_that("as_POSIXct_utc", {
  exp <- "2022-03-02 19:27:35"
  got <- as_POSIXct_utc("2022-03-02 19:27:35")
  expect_equal(exp, as.character(got))
})

test_that("convert to date", {
  dates <- as.Date("2015-02-07") + -10:10
  origin <- 25569L
  n <- as.integer(dates) + origin

  expect_equal(convertToDate(n), dates)

  earlyDate <- as.Date("1900-01-03")
  serialDate <- 3
  expect_equal(convertToDate(serialDate), earlyDate)

})


test_that("convert to datetime", {
  x <- 43037 + 2 / 1440
  expect_equal(object = convertToDateTime(x, tx = Sys.timezone()), expected = as.POSIXct("2017-10-29 00:02:00", tz = Sys.timezone()))

  x <- 43037 + 2 / 1440 + 1 / 86400
  expect_equal(object = convertToDateTime(x, tx = Sys.timezone()), expected = as.POSIXct("2017-10-29 00:02:01", tz = Sys.timezone()))

  x <- 43037 + 2.50 / 1440
  expect_equal(object = convertToDateTime(x, tx = Sys.timezone()), expected = as.POSIXct("2017-10-29 00:02:30", tz = Sys.timezone()))

  x <- 43037 + 2 / 1440 + 12.12 / 86400
  x_datetime <- convertToDateTime(x, tx = "UTC")
  attr(x_datetime, "tzone") <- "UTC"
})


test_that("read.xlsx detectDates", {
  
  xlsxFile <- system.file("extdata", "gh_issue_288.xlsx", package = "openxlsx")
  
  ref_dat <- data.frame(Date = c(as.Date(c("2021-10-20", "2021-11-03"))))
  ref_num <- data.frame(Date = c(44489.4, 44503.0))
  
  tst_dat <- read.xlsx(xlsxFile, detectDates = TRUE)
  tst_num <- read.xlsx(xlsxFile, detectDates = FALSE)
  
  expect_equal(ref_dat, tst_dat)
  expect_equal(ref_num, tst_num)
  
})
