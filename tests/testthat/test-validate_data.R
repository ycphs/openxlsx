context("Data validation")

# Basic test workbook ==========================================================

wb <- createWorkbook()
addWorksheet(wb, "Sheet 1")
addWorksheet(wb, "Sheet 2")
writeDataTable(wb, sheet = 1, x = iris[1:30, ])
writeData(wb, sheet = 2, x = sample(iris$Sepal.Length, 10))

# Unit tests ===================================================================

test_that("Data validation for lists is performed without warnings", {
  expect_invisible(
    dataValidation(
      wb, 1, col = 1, rows = 2:31, type = "list", value = "'Sheet 2'!$A$1:$A$10"
    )
  )
})
