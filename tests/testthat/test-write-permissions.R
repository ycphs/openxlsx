
context("error without write permissions")

test_that("test failed write errors for saveWorkbook", {
  
  skip_on_cran()
  skip_on_ci()
  
  temp_file <- tempfile()
  file.create(temp_file)
  Sys.chmod(path = temp_file, mode = "444")

  wb <- createWorkbook()
  addWorksheet(wb, "name")

  expect_warning(write.xlsx(
    x = cars, file = temp_file, overwrite = TRUE
  ),
  regexp = "Permission denied"
  )

  unlink(temp_file, recursive = TRUE, force = TRUE)
})
