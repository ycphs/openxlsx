context("Date detection: format-code length check")

test_that("short date-like format codes are not detected as dates", {
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)

  wb <- createWorkbook()
  addWorksheet(wb, "Sheet1")
  writeData(wb, "Sheet1", 44927L, colNames = FALSE)
  addStyle(wb, "Sheet1", createStyle(numFmt = "m"), rows = 1, cols = 1, stack = FALSE)
  saveWorkbook(wb, tmp, overwrite = TRUE)

  expect_false(inherits(
    read.xlsx(tmp, sheet = 1, colNames = FALSE, detectDates = TRUE)[[1]],
    "Date"
  ))
  expect_false(inherits(
    read.xlsx(wb, sheet = 1, colNames = FALSE, detectDates = TRUE)[[1]],
    "Date"
  ))
})
