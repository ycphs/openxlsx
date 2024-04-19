test_that("can read default sizes from new workbook", {
  wb <- createWorkbook()
  
  expect_equal(getWindowSize(wb), 
               list(xWindow = "0", yWindow = "0", windowWidth = "13125", windowHeight = "6105"))
})

test_that("can change sizes in new workbook", {
  wb <- createWorkbook()
  setWindowSize(wb, xWindow = 100, yWindow = 200, windowWidth = 20000, windowHeight = 15000)
  
  expect_equal(getWindowSize(wb), 
               list(xWindow = "100", yWindow = "200", windowWidth = "20000", windowHeight = "15000"))
})

test_that("can change only a few parameters", {
  wb <- createWorkbook()
  setWindowSize(wb, yWindow = 300, windowWidth = 12000) # others are the default
  
  expect_equal(getWindowSize(wb), 
               list(xWindow = "0", yWindow = "300", windowWidth = "12000", windowHeight = "6105"))
})

test_that("window size and position must be coercible to integers and non negative", {
  wb <- createWorkbook()
  expect_error(setWindowSize(wb, xWindow = -300))
  expect_error(setWindowSize(wb, yWindow = -300))
  expect_error(setWindowSize(wb, windowWidth = -300))
  expect_error(setWindowSize(wb, windowHeight = -300))
  expect_error(expect_warning(setWindowSize(wb, xWindow = "string")))
  expect_error(expect_warning(setWindowSize(wb, yWindow = "string")))
  expect_error(expect_warning(setWindowSize(wb, windowWidth = "string")))
  expect_error(expect_warning(setWindowSize(wb, windowHeight = "string")))
  
  setWindowSize(wb, xWindow = "500")
  expect_equal(getWindowSize(wb)$xWindow, "500")
})

test_that("saving and loading preserve window sizes", {
  wb <- createWorkbook()
  setWindowSize(wb, xWindow = 400, yWindow = 500, windowWidth = 1000, windowHeight = 5000)
  addWorksheet(wb, "sheet1")
  tempfile <- tempfile()
  saveWorkbook(wb, tempfile)
  wb2 <- loadWorkbook(tempfile)
  
  expect_equal(getWindowSize(wb2),
               list(xWindow = "400", yWindow = "500", windowWidth = "1000", windowHeight = "5000"))
  
  setWindowSize(wb2, xWindow = 700, yWindow = 900, windowWidth = 900, windowHeight = 4000)
  
  tempfile2 <- tempfile()
  saveWorkbook(wb2, tempfile2)
  wb3 <- loadWorkbook(tempfile2)
  
  expect_equal(getWindowSize(wb3),
               list(xWindow = "700", yWindow = "900", windowWidth = "900", windowHeight = "4000"))
})

