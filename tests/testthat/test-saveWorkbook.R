
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


# test for hyperrefs
test_that("creating hyperlinks", {
  
  # prepare a file
  tempFile <- temp_xlsx()
  wb <- createWorkbook()
  sheet <- "test"
  addWorksheet(wb, sheet)
  img <- "D:/somepath/somepicture.png"
  
  # warning: col and row provided, but not required
  expect_warning(
    linkString <- makeHyperlinkString(col = 1, row = 4,
                                      text = "test.png", file = img))
  
  linkString2 <- makeHyperlinkString(text = "test.png", file = img)
  
  # col and row not needed
  expect_equal(linkString, linkString2)
  
  # write file without errors
  writeFormula(wb, sheet, x = linkString, startCol = 1, startRow = 1)
  expect_silent(saveWorkbook(wb, tempFile, overwrite = TRUE))
  
  # TODO: add a check that the written xlsx file contains linkString
  
})
