test_that("writeDataTable() warns about Excel limitations", {

  dt <- data.frame(var = c(1))
  header_255 <- paste(sample(tolower(LETTERS), 255, TRUE), collapse = "")
  header_256 <- paste(sample(tolower(LETTERS), 256, TRUE), collapse = "")

  wb <- createWorkbook()
  addWorksheet(wb, "sheetA")
  addWorksheet(wb, "sheetB")

  expect_warning({
    colnames(dt) <- header_255
    writeDataTable(wb, sheet = "sheetA", x = dt)
    },
    regexp = NA
  )

  expect_warning({
    colnames(dt) <- header_256
    writeDataTable(wb, sheet = "sheetB", x = dt)
    },
    regexp = "Column name exceeds 255 chars, possible incompatibility with MS Excel. Index: 1"
  )
})