

context("Reading from workbook is identical to reading from file")



test_that("Reading from loaded workbook DF", {
 
  wb <- createWorkbook()
  for(i in 1:4)
    addWorksheet(wb, sprintf('Sheet %s', i))
  
  writeData(wb, sheet = 1, x = mtcars, colNames = TRUE, rowNames = TRUE, startRow = 10, startCol = 5, borders = "all")
  writeData(wb, sheet = 2, x = mtcars, colNames = TRUE, rowNames = FALSE, startRow = 10, startCol = 5, borders = "rows")
  writeData(wb, sheet = 3, x = mtcars, colNames = FALSE, rowNames = TRUE, startRow = 2, startCol = 2, borders = "columns")
  writeData(wb, sheet = 4, x = mtcars, colNames = FALSE, rowNames = FALSE, startRow = 12, startCol = 1, borders = "surrounding")
  
  tempFile <- file.path(tempdir(), "temp.xlsx")
  saveWorkbook(wb, tempFile, overwrite = TRUE)
  
  wb <- loadWorkbook(tempFile)
  
  ## colNames = TRUE, rowNames = TRUE
  x <- read.xlsx(wb, 1, colNames = TRUE, rowNames = TRUE, asdatatable = FALSE)
  # attributes(x)$row.names<-as.integer(attributes(x)$row.names)
  expect_equal(object = x, expected = data.frame(mtcars), check.attributes = TRUE)
  

  
  
  ## colNames = TRUE, rowNames = FALSE
  x <- read.xlsx(wb, sheet = 2, colNames = TRUE, rowNames = FALSE, asdatatable = TRUE)
  expect_equal(object = mtcars, expected = x, check.attributes = FALSE)
  
  expect_equal(object = colnames(mtcars), expected = colnames(x), check.attributes = FALSE)
  
  
  ## colNames = FALSE, rowNames = TRUE
  x <- read.xlsx(wb, sheet = 3, colNames = FALSE, rowNames = TRUE, asdatatable = FALSE)

  expect_equal(object = data.frame(mtcars), expected = x, check.attributes = FALSE)
  expect_equal(object = rownames(mtcars), expected = rownames(x))
  
  
  ## colNames = FALSE, rowNames = FALSE
  x <- read.xlsx(wb, sheet = 4, colNames = FALSE, rowNames = FALSE, asdatatable = FALSE)  
  attributes(x)$row.names<-as.integer(attributes(x)$row.names)
  expect_equal(object = data.frame(mtcars), expected = x, check.attributes = FALSE)
  
  unlink(tempFile, recursive = TRUE, force = TRUE)
  
})

test_that("Reading from loaded workbook DT", {
  library(data.table)
  wb <- createWorkbook()
  for(i in 1:4)
    addWorksheet(wb, sprintf('Sheet %s', i))
  
  writeData(wb, sheet = 1, x = mtcars, colNames = TRUE, rowNames = TRUE, startRow = 10, startCol = 5, borders = "all")
  writeData(wb, sheet = 2, x = mtcars, colNames = TRUE, rowNames = FALSE, startRow = 10, startCol = 5, borders = "rows")
  writeData(wb, sheet = 3, x = mtcars, colNames = FALSE, rowNames = TRUE, startRow = 2, startCol = 2, borders = "columns")
  writeData(wb, sheet = 4, x = mtcars, colNames = FALSE, rowNames = FALSE, startRow = 12, startCol = 1, borders = "surrounding")
  
  tempFile <- file.path(tempdir(), "temp.xlsx")
  
  saveWorkbook(wb, tempFile, overwrite = TRUE)
  mtc<-mtcars
  setDT(mtc)
  wb <- loadWorkbook(tempFile)
  
  ## colNames = TRUE, rowNames = TRUE
  x <- read.xlsx(wb, 1, colNames = TRUE, rowNames = TRUE, asdatatable = TRUE)
  expect_equal(object = mtc, expected = x, check.attributes = TRUE)
  
  
  ## colNames = TRUE, rowNames = FALSE
  x <- read.xlsx(wb, sheet = 2, colNames = TRUE, rowNames = FALSE, asdatatable = TRUE)
  expect_equal(object = mtc, expected = x, check.attributes = FALSE)
  expect_equal(object = colnames(mtc), expected = colnames(x), check.attributes = FALSE)
  
  
  ## colNames = FALSE, rowNames = TRUE
  x <- read.xlsx(wb, sheet = 3, colNames = FALSE, rowNames = TRUE, asdatatable = TRUE)
  expect_equal(object = mtc, expected = x, check.attributes = FALSE)
  expect_equal(object = rownames(mtc), expected = rownames(x))
  
  
  ## colNames = FALSE, rowNames = FALSE
  x <- read.xlsx(wb, sheet = 4, colNames = FALSE, rowNames = FALSE, asdatatable = TRUE)  
  expect_equal(object = mtc, expected = x, check.attributes = FALSE)
  
  unlink(tempFile, recursive = TRUE, force = TRUE)
  
})
