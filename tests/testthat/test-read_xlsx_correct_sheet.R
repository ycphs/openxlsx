

context("Read xlsx")


test_that("read.xlsx correct sheet", {
  library(data.table)
  fl <- system.file("readTest.xlsx", package = "openxlsx")
  sheet_names <- getSheetNames(file = fl)
  
  expected_sheet_names <- c("Sheet1", "Sheet2", "Sheet 3",
                            "Sheet 4", "Sheet 5", "Sheet 6", 
                            "1", "11", "111", "1111", "11111", "111111")
  
  
  expect_equal(object = sheet_names, expected = expected_sheet_names) 
  
  expect_equal(read.xlsx(xlsxFile = fl, sheet = 7), data.table(x = 1))
  expect_equal(read.xlsx(xlsxFile = fl, sheet = 8), data.table(x = 11))
  expect_equal(read.xlsx(xlsxFile = fl, sheet = 9), data.table(x = 111))
  expect_equal(read.xlsx(xlsxFile = fl, sheet = 10), data.table(x = 1111))
  expect_equal(read.xlsx(xlsxFile = fl, sheet = 11), data.table(x = 11111)) 
  expect_equal(read.xlsx(xlsxFile = fl, sheet = 12), data.table(x = 111111)) 
  
  expect_equal(read.xlsx(xlsxFile = fl, sheet = "1"), data.table(x = 1))
  expect_equal(read.xlsx(xlsxFile = fl, sheet = "11"), data.table(x = 11))
  expect_equal(read.xlsx(xlsxFile = fl, sheet = "111"), data.table(x = 111))
  expect_equal(read.xlsx(xlsxFile = fl, sheet = "1111"), data.table(x = 1111))
  expect_equal(read.xlsx(xlsxFile = fl, sheet = "11111"), data.table(x = 11111)) 
  expect_equal(read.xlsx(xlsxFile = fl, sheet = "111111"), data.table(x = 111111)) 
  
  expect_equal(read.xlsx(xlsxFile = fl, asdatatable = FALSE, sheet = 7), data.frame(x = 1))
  expect_equal(read.xlsx(xlsxFile = fl, asdatatable = FALSE, sheet = 8), data.frame(x = 11))
  expect_equal(read.xlsx(xlsxFile = fl, asdatatable = FALSE, sheet = 9), data.frame(x = 111))
  expect_equal(read.xlsx(xlsxFile = fl, asdatatable = FALSE, sheet = 10), data.frame(x = 1111))
  expect_equal(read.xlsx(xlsxFile = fl, asdatatable = FALSE, sheet = 11), data.frame(x = 11111)) 
  expect_equal(read.xlsx(xlsxFile = fl, asdatatable = FALSE, sheet = 12), data.frame(x = 111111)) 
  
  expect_equal(read.xlsx(xlsxFile = fl, asdatatable = FALSE, sheet = "1"), data.frame(x = 1))
  expect_equal(read.xlsx(xlsxFile = fl, asdatatable = FALSE, sheet = "11"), data.frame(x = 11))
  expect_equal(read.xlsx(xlsxFile = fl, asdatatable = FALSE, sheet = "111"), data.frame(x = 111))
  expect_equal(read.xlsx(xlsxFile = fl, asdatatable = FALSE, sheet = "1111"), data.frame(x = 1111))
  expect_equal(read.xlsx(xlsxFile = fl, asdatatable = FALSE, sheet = "11111"), data.frame(x = 11111)) 
  expect_equal(read.xlsx(xlsxFile = fl, asdatatable = FALSE, sheet = "111111"), data.frame(x = 111111)) 
  
  
  
})

