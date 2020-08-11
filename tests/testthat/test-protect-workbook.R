

context("Protection")


test_that("Protect Workbook", {
  wb <- createWorkbook()
  addWorksheet(wb, "s1")

  wb$protectWorkbook(password = "abcdefghij")

  expect_true(wb$workbook$workbookProtection == "<workbookProtection workbookPassword=\"FEF1\"/>")

  wb$protectWorkbook(protect = FALSE, password = "abcdefghij", lockStructure = TRUE, lockWindows = TRUE)
  expect_true(wb$workbook$workbookProtection == "")
})

test_that("Reading protected Workbook", {
  tmp_file <- file.path(tempdir(), "xlsx_read_protectedwb.xlsx")

  wb <- createWorkbook()
  addWorksheet(wb, "s1")
  protectWorkbook(wb, password = "abcdefghij")
  saveWorkbook(wb, tmp_file, overwrite = TRUE)

  wb2 <- loadWorkbook(file = tmp_file)
  # Check that the order of teh sub-elements is preserved
  expect_equal(names(wb2$workbook), names(wb$workbook))

  unlink(tmp_file, recursive = TRUE, force = TRUE)
})
