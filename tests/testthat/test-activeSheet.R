
context("active Sheet ")


test_that("get and set active sheet of a workbook", {
  
  tempFile1 <- temp_xlsx("temp1")
  tempFile2 <- temp_xlsx("temp2")
  tempFile3 <- temp_xlsx("temp3")
  wbook <- createWorkbook()
  addWorksheet(wbook, sheetName = "S1")
  addWorksheet(wbook, sheetName = "S2")
  addWorksheet(wbook, sheetName = "S3")
  
  
  saveWorkbook(wbook,tempFile1)
  # default value is the first sheet active
  expect_equal(activeSheet(wbook),1)
  expect_equal(activeSheet(wbook),loadWorkbook(tempFile1)$ActiveSheet)
  
  activeSheet(wbook) <- 1 ## active sheet S1
  saveWorkbook(wbook,tempFile2)
  expect_equal(activeSheet(wbook),1)
  expect_equal(activeSheet(wbook),loadWorkbook(tempFile2)$ActiveSheet)
  activeSheet(wbook) <- "S2" ## active sheet S2
  saveWorkbook(wbook,tempFile3)
  expect(activeSheet(wbook),2)
  expect_equal(activeSheet(wbook),loadWorkbook(tempFile3)$ActiveSheet)
  
  unlink(tempFile1, recursive = TRUE, force = TRUE)
  unlink(tempFile2, recursive = TRUE, force = TRUE)
  unlink(tempFile3, recursive = TRUE, force = TRUE)
})
