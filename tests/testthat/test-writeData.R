test_that("writeData() forces evaluation of x (#264)", {
  wbfile <- temp_xlsx()
  op <- options(stringsAsFactors = FALSE)
  
  x <- format(123.4)
  df <- data.frame(d = format(123.4))
  df2 <- data.frame(e = x)
  
  wb <- createWorkbook()
  addWorksheet(wb, "sheet")
  writeData(wb, "sheet", startCol = 1, data.frame(a = format(123.4)))
  writeData(wb, "sheet", startCol = 2, data.frame(b = as.character(123.4)))
  writeData(wb, "sheet", startCol = 3, data.frame(c = "123.4"))
  writeData(wb, "sheet", startCol = 4, df)
  writeData(wb, "sheet", startCol = 5, df2)
  
  saveWorkbook(wb, wbfile)
  out <- read.xlsx(wbfile)
  
  # Possibly overkill
  
  with(out, {
    expect_identical(a, b)
    expect_identical(a, c)
    expect_identical(a, d)
    expect_identical(a, e)
    expect_identical(b, c)
    expect_identical(b, d)
    expect_identical(b, e)
    expect_identical(c, d)
    expect_identical(c, e)
    expect_identical(d, e)
  })
  
  options(op)
  file.remove(wbfile)
})

test_that("colnames with NA are appropriately handled [292]", {
  x <- data.frame(a = 1, b = 2)
  colnames(x) <- c("a", NA)
  wbfile <- temp_xlsx()
  wb <- buildWorkbook(x)
  wb$worksheets[[1]]$sheet_data
  saveWorkbook(wb, wbfile)
  
  # what could I look for 
  # openXL(wb)
  # openXL(wbfile)
  # read.xlsx(wbfile)
})
