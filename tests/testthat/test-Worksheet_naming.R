
context("Worksheet naming")

test_that("Worksheet names", {
  
  ### test for names without special character
  wb <- createWorkbook()
  sheetname <- "test"
  addWorksheet(wb, sheetname)
  
  expect_equal(sheetname,names(wb))
  
  ### test for names with &
  
  wb <- createWorkbook()
  sheetname <- "S&P 500"
  addWorksheet(wb, sheetname)
  
  expect_equal(sheetname,names(wb))
  expect_equal("S&amp;P 500",wb$sheet_names)
  ### test for names with <
  
  wb <- createWorkbook()
  sheetname <- "<24 h"
  addWorksheet(wb, sheetname)
  
  expect_equal(sheetname,names(wb))
  expect_equal("&lt;24 h",wb$sheet_names)
  ### test for names with >
  
  wb <- createWorkbook()
  sheetname <- ">24 h"
  addWorksheet(wb, sheetname)
  
  expect_equal(sheetname,names(wb))
  expect_equal("&gt;24 h",wb$sheet_names)
  
  ### test for names with "
  
  wb <- createWorkbook()
  sheetname <- 'test "A"'
  addWorksheet(wb, sheetname)
  
  expect_equal(sheetname,names(wb))
  expect_equal("test &quot;A&quot;",wb$sheet_names)
})


test_that("test for illegal characters", {
  ### 
  
  wb <- createWorkbook()
  x <- data.frame(a = 1, b = 2)
  
  addWorksheet(wb, "Test")
  for(i in c("[", "]", "*", "/",  "?", ":")){
    sheetname <- paste0('test_',i)
    y <- list(a = x, sheetname = x, c = x)
    expect_error(addWorksheet(wb, sheetname))
    expect_error(cloneWorksheet(wb, "Test", sheetname))
    expect_error(renameWorksheet(wb, "Test", sheetname))
    expect_error(parse(names(wb)<- sheetname))
    expect_error(parse(buildWorkbook(y, asTable = TRUE)))
  }
  
})
  
  
  
  
  

