
context("save workbook")


test_that("test return values for saveWorkbook", {
  
  tempFile <- temp_xlsx()
  wb<-createWorkbook()
  addWorksheet(wb,"name")
  expect_true( saveWorkbook(wb,tempFile,returnValue = TRUE))
  
  expect_error( saveWorkbook(wb,tempFile,returnValue = TRUE))
  
  
  expect_invisible( saveWorkbook(wb,tempFile,returnValue = FALSE ,overwrite = TRUE))
  unlink(tempFile, recursive = TRUE, force = TRUE)
  
}
)

# regression test for a typo
test_that("regression test for #248", {
  
  # Basic data frame
  df <- data.frame(number = 1:3, percent = 4:6/100)
  tempFile <- temp_xlsx()
  
  # no formatting
  expect_silent(write.xlsx(df, tempFile, borders = "columns", overwrite = TRUE))
  
  # Change column class to percentage
  class(df$percent) <- "percentage"
  expect_silent(write.xlsx(df, tempFile, borders = "columns", overwrite = TRUE))
})
