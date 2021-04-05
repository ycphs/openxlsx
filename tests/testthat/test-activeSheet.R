
context("active Sheet ")


test_that("get and set active sheet of a workbook", {
  wbook <- createWorkbook()
  addWorksheet(wbook, sheetName = "S1")
  addWorksheet(wbook, sheetName = "S2")
  addWorksheet(wbook, sheetName = "S3")
  
   # default value is the first sheet active
  expect_equal(activeSheet(wbook),1)
  activeSheet(wbook) <- 1 ## active sheet S1
  expect_equal(activeSheet(wbook),1)
  activeSheet(wbook) <- "S2" ## active sheet S2
  expect(activeSheet(wbook),2)
  

})
