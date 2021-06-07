




context("Fill Merged Cells")



test_that("fill merged cells", {
  wb <- createWorkbook()
  addWorksheet(wb, sheetName = "sheet1")
  writeData(wb = wb, sheet = 1, x = data.frame("A" = 1, "B" = 2))
  writeData(wb = wb, sheet = 1, x = 2, startRow = 2, startCol = 2)
  writeData(wb = wb, sheet = 1, x = 3, startRow = 2, startCol = 3)
  writeData(wb = wb, sheet = 1, x = 4, startRow = 2, startCol = 4)
  writeData(wb = wb, sheet = 1, x = t(matrix(1:4, 4, 4)), startRow = 3, startCol = 1, colNames = FALSE)

  mergeCells(wb = wb, sheet = 1, cols = 2:4, rows = 1)
  mergeCells(wb = wb, sheet = 1, cols = 2:4, rows = 3)
  mergeCells(wb = wb, sheet = 1, cols = 2:4, rows = 4)
  mergeCells(wb = wb, sheet = 1, cols = 2:4, rows = 5)

  tmp_file <- temp_xlsx()
  saveWorkbook(wb = wb, file = tmp_file, overwrite = TRUE)

  expect_equal(names(read.xlsx(tmp_file, fillMergedCells = FALSE)), c("A", "B", "X3", "X4"))
  expect_equal(names(read.xlsx(tmp_file, fillMergedCells = TRUE)), c("A", "B", "B", "B"))
  
  r1 <- data.frame("A" = rep(1, 5), "B" = rep(2, 5), "X3" = rep(3,5), "X4" = rep(4, 5))
  r2 <- data.frame("A" = rep(1, 5), "B" = rep(2, 5), "B1" = c(3,2,2,2,3), "B2" = c(4,2,2,2,4))
  names(r2) <- c("A", "B", "B", "B")
  
  r2_1 <- r2[1:5, 1:3]
  names(r2_1) <- c("A", "B", "B")
  
  expect_equal(read.xlsx(tmp_file, fillMergedCells = FALSE), r1)
  expect_equal(read.xlsx(tmp_file, fillMergedCells = TRUE), r2)
  
  expect_equal( read.xlsx(tmp_file, cols = 1:3, fillMergedCells = TRUE), r2_1)
  expect_equal( read.xlsx(tmp_file, rows = 1:3, fillMergedCells = TRUE), r2[1:2, ])
  expect_equal( read.xlsx(tmp_file, cols = 1:3, rows = 1:4, fillMergedCells = TRUE), r2_1[1:3,])
  
})
