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

test_that("row.name and col.name are deprecated", {
  x <- data.frame(a = 1)
  
  expect_warning(
    buildWorkbook(x, file = temp_xlsx(), row.names = TRUE, overwrite = TRUE),
    "Please use 'rowNames' instead of 'row.names'"
  )
  
  expect_warning(
    buildWorkbook(x, file = temp_xlsx(), row.names = TRUE, overwrite = TRUE, asTable = TRUE),
    "Please use 'rowNames' instead of 'row.names'"
  )
  
  expect_warning(
    buildWorkbook(x, file = temp_xlsx(), col.names = TRUE, overwrite = TRUE),
    "Please use 'colNames' instead of 'col.names'"
  )
  
  expect_warning(
    buildWorkbook(x, file = temp_xlsx(), col.names = TRUE, overwrite = TRUE, asTable = TRUE),
    "Please use 'colNames' instead of 'col.names'"
  )
})
