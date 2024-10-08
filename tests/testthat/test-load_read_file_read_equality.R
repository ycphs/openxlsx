

context("Reading from workbook is identical to reading from file")

test_that("Reading from loaded workbook", {
  wb <- createWorkbook()
  for (i in 1:4) {
    addWorksheet(wb, sprintf("Sheet %s", i))
  }

  writeData(wb, sheet = 1, x = mtcars, colNames = TRUE, rowNames = TRUE, startRow = 10, startCol = 5, borders = "all")
  writeData(wb, sheet = 2, x = mtcars, colNames = TRUE, rowNames = FALSE, startRow = 10, startCol = 5, borders = "rows")
  writeData(wb, sheet = 3, x = mtcars, colNames = FALSE, rowNames = TRUE, startRow = 2, startCol = 2, borders = "columns")
  writeData(wb, sheet = 4, x = mtcars, colNames = FALSE, rowNames = FALSE, startRow = 12, startCol = 1, borders = "surrounding")

  tempFile <- temp_xlsx()
  saveWorkbook(wb, tempFile, overwrite = TRUE)

  wb <- loadWorkbook(tempFile)

  ## colNames = TRUE, rowNames = TRUE
  x <- read.xlsx(wb, 1, colNames = TRUE, rowNames = TRUE)
  expect_equal(object = mtcars, expected = x, check.attributes = TRUE)


  ## colNames = TRUE, rowNames = FALSE
  x <- read.xlsx(wb, sheet = 2, colNames = TRUE, rowNames = FALSE)
  expect_equal(object = mtcars, expected = x, check.attributes = FALSE)
  expect_equal(object = colnames(mtcars), expected = colnames(x), check.attributes = FALSE)


  ## colNames = FALSE, rowNames = TRUE
  x <- read.xlsx(wb, sheet = 3, colNames = FALSE, rowNames = TRUE)
  expect_equal(object = mtcars, expected = x, check.attributes = FALSE)
  expect_equal(object = rownames(mtcars), expected = rownames(x))


  ## colNames = FALSE, rowNames = FALSE
  x <- read.xlsx(wb, sheet = 4, colNames = FALSE, rowNames = FALSE)
  expect_equal(object = mtcars, expected = x, check.attributes = FALSE)

  unlink(tempFile, recursive = TRUE, force = TRUE)
})


test_that("na.convert = TRUE", {
  
  fl <- system.file("extdata", "na_convert.xlsx", package = "openxlsx")
  
  
  # optional return empty cell as "<si><t/></si>"
  wb1 <- loadWorkbook(fl, na.convert = FALSE)
  tmp_no_na <- temp_xlsx("na_convert_false")
  saveWorkbook(wb1, tmp_no_na)
  
  # default return emtpy cell as "NA" string
  wb2 <- loadWorkbook(fl, na.convert = TRUE)
  tmp_has_na <- temp_xlsx("na_convert_true")
  saveWorkbook(wb2, tmp_has_na)
  
  # # FIXME both return identical output
  # read.xlsx(wb1)
  # read.xlsx(wb2)
  
  # # for visual comparison
  # openXL(tmp_has_na)
  # openXL(tmp_no_na)
  
  ## new option
  # check value of last cells: index position of shared strings
  exp <- c("1", "0", "0", "0", "0", "0")
  got <- tail(wb1$worksheets[[1]]$sheet_data$v)
  expect_equal(exp, got)
  
  # check value of shared string: index position 0
  exp <- "<si><t/></si>"
  got <- wb1$sharedStrings[[1]]
  expect_equal(exp, got)
  
  
  ## default
  # check value of last cells: index position of shared strings
  exp <- c("1", "0", "0", "0", "0", "0")
  got <- tail(wb2$worksheets[[1]]$sheet_data$v)
  expect_equal(exp, got)
  
  # check value of shared string: index position 0
  exp <- "<si><t>NA</t></si>"
  got <- wb2$sharedStrings[[1]]
  expect_equal(exp, got)
  
})
