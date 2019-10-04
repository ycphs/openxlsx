



context("Encoding Tests")



test_that("Write read encoding equality DT", {
  library(data.table)
  tempFile <- file.path(tempdir(), "temp.xlsx")
  
  wb <- createWorkbook()
  for (i in 1:4)
    addWorksheet(wb, sprintf('Sheet %s', i))
  
  df <- data.table("X" = c("测试", "一下"), stringsAsFactors = FALSE)
  writeDataTable(wb, sheet = 1, x = df)
  
  saveWorkbook(wb, tempFile, overwrite = TRUE)
  
  x <- read.xlsx(tempFile, asdatatable = TRUE)
  expect_equal(x, df)
  
  x <- read.xlsx(wb, asdatatable = TRUE)
  expect_equal(x, df)
  
  ## reload
  wb <- loadWorkbook(tempFile)
  
  x <- read.xlsx(wb)
  expect_equal(x, df)
  
  saveWorkbook(wb, tempFile, overwrite = TRUE)
  x <- read.xlsx(tempFile)
  expect_equal(x, df)
  
  unlink(tempFile, recursive = TRUE, force = TRUE)
  rm(wb)
  
})


test_that("Write read encoding equality DF", {
  tempFile <- file.path(tempdir(), "temp.xlsx")
  
  wb <- createWorkbook()
  for (i in 1:4)
    addWorksheet(wb, sprintf('Sheet %s', i))
  
  df <- data.frame("X" = c("测试", "一下"), stringsAsFactors = FALSE)
  writeDataTable(wb, sheet = 1, x = df)
  
  saveWorkbook(wb, tempFile, overwrite = TRUE)
  
  x <- read.xlsx(tempFile, asdatatable = FALSE)
  expect_equal(x, df)
  
  x <- read.xlsx(wb, asdatatable = FALSE)
  expect_equal(x, df)
  
  ## reload
  wb <- loadWorkbook(tempFile)
  
  x <- read.xlsx(wb, asdatatable = FALSE)
  expect_equal(x, df)
  
  saveWorkbook(wb, tempFile, overwrite = TRUE)
  x <- read.xlsx(tempFile, asdatatable = FALSE)
  expect_equal(x, df)
  
  unlink(tempFile, recursive = TRUE, force = TRUE)
  rm(wb)
  
})
