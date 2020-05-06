

context("clone Worksheet")


test_that("clone Worksheet with data", {
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  writeData(wb, "Sheet 1", 1)
  cloneWorksheet(wb, "Sheet 2", clonedSheet = "Sheet 1")


  file_name <- system.file("extdata", "cloneWorksheetExample.xlsx", package = "openxlsx")
  refwb <- loadWorkbook(file = file_name)

  expect_equal(sheets(wb), sheets(refwb))
  expect_equal(worksheetOrder(wb), worksheetOrder(refwb))
})

test_that("clone empty Worksheet", {
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")

  cloneWorksheet(wb, "Sheet 2", clonedSheet = "Sheet 1")


  file_name <- system.file("extdata", "cloneEmptyWorksheetExample.xlsx", package = "openxlsx")
  refwb <- loadWorkbook(file = file_name)

  expect_equal(sheets(wb), sheets(refwb))
  expect_equal(worksheetOrder(wb), worksheetOrder(refwb))
})
