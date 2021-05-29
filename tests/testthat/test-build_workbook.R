test_that("buildWorkbook() accepts tableName [187]", {
  x <- data.frame(a = 1, b = 2)
  
  # default name
  wb <- buildWorkbook(x, asTable = TRUE)
  expect_equal(attr(wb$tables, "tableName"), "Table3")
  
  # define 1/2 table name
  wb <- buildWorkbook(x, asTable = TRUE, tableName = "table_x")
  expect_equal(attr(wb$tables, "tableName"), "table_x")
  
  # define 2/2 table names
  wb <- buildWorkbook(list(x, x), asTable = TRUE, tableName = c("table_x", "table_y"))
  expect_equal(attr(wb$tables, "tableName"), c("table_x", "table_y"))
  
  # try to define 1/2 table names
  expect_error(buildWorkbook(list(x, x), asTable = TRUE, tableName = "table_x"))
})
