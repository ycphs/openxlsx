

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
