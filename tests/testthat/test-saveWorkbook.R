
context("save workbook")


test_that("test return values for saveWorkbook", {
  
  tempFile <- file.path(tempdir(), "temp.xlsx")
  wb<-createWorkbook()
  addWorksheet(wb,"name")
  expect_true( saveWorkbook(wb,tempFile,returnValue = TRUE))
  
  expect_error( saveWorkbook(wb,tempFile,returnValue = TRUE))
  
  
  expect_invisible( saveWorkbook(wb,tempFile,returnValue = FALSE ,overwrite = TRUE))
  unlink(tempFile, recursive = TRUE, force = TRUE)
  
}
)