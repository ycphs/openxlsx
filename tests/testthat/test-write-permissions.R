
context("error without write permissions")

test_that("test failed write errors for saveWorkbook", {
  tempFile <- fs::file_temp(ext = "xlsx")
  fs::file_create(tempFile, mode = "a=rx")

  wb <- createWorkbook()
  addWorksheet(wb, "name")

  expect_warning(write.xlsx(
    x = cars, file = tempFile, overwrite = TRUE
  ),
  regexp = "Permission denied"
  )

  unlink(tempFile, recursive = TRUE, force = TRUE)
})
