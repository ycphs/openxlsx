

context("v3.0.0 Bug Fixes")



test_that("read.xlsx bug fixes", {
  library(data.table)
  file <- system.file("readTest.xlsx", package = "openxlsx")
  df <- read.xlsx(file, sheet = 1, rows = 1:2, cols = 1, asdatatable = FALSE)
  expect_equal(df, data.frame("Var1" = TRUE))
  
  df <- read.xlsx(file, sheet = 1, rows = 1, cols = 1, colNames = FALSE, asdatatable = FALSE)
  expect_equal(df, data.frame("X1" = "Var1", stringsAsFactors = FALSE))
  
  dt <- read.xlsx(file, sheet = 1, rows = 1:2, cols = 1, asdatatable = TRUE)
  expect_equal(dt, data.table("Var1" = TRUE))
  
  dt <- read.xlsx(file, sheet = 1, rows = 1, cols = 1, colNames = FALSE, asdatatable = TRUE)
  expect_equal(dt, data.table("X1" = "Var1", stringsAsFactors = FALSE))
  
  
})

