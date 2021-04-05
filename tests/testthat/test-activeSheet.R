
context("active Sheet ")


test_that("get and set active sheet of a workbook", {
  wb <- createWorkbook()
  addWorksheet(wb, sheetName = "S1")
  addWorksheet(wb, sheetName = "S2")
  addWorksheet(wb, sheetName = "S3")
  
  activeSheet(wb) # default value is the first sheet active
  activeSheet(wb) <- 1 ## active sheet S1
  expect(activeSheet(wb),1)
  activeSheet(wb) <- "S2" ## active sheet S2
  expect(activeSheet(wb),2)
  

})
