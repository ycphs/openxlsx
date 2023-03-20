
context("Test wrappers")

test_that("int2col and col2int", {
  
  nums <- 2:27
  
  chrs <- c("B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",  "N",
            "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA")
  
  expect_equal(chrs, int2col(nums))
  expect_equal(nums, col2int(chrs))
  
})

test_that("deleteDataColumn", {
  wb <- createWorkbook()
  addWorksheet(wb, "tester")

  for (i in seq(5)) {
    mat <- data.frame(x = rep(paste0(int2col(i), i), 10))
    writeData(wb, sheet = 1, startRow = 1, startCol = i, mat)
    writeFormula(wb, sheet = 1, startRow = 12, startCol = i,
                 x = sprintf("=COUNTA(%s2:%s11)", int2col(i), int2col(i)))
  }
  deleteDataColumn(wb, 1, col = 3)
  
  expect_equal(read.xlsx(wb),
               data.frame(x = rep("A1", 10), x = "B2", x = "D4", x = "E5", # no C3!
                          check.names = FALSE))
  
  deleteDataColumn(wb, 1, col = 2)
  expect_equal(read.xlsx(wb),
               data.frame(x = rep("A1", 10), x = "D4", x = "E5", # no B2!
                          check.names = FALSE))

  deleteDataColumn(wb, 1, col = 1)
  expect_equal(read.xlsx(wb),
               data.frame(x = rep("D4", 10), x = "E5", # no A1!
                          check.names = FALSE))
  
  # works with formula as well!
  wb <- createWorkbook()
  addWorksheet(wb, "tester")
  writeData(wb, sheet = 1, startRow = 1, startCol = 1,
            x = matrix(c(1, 1), ncol = 1), colNames = FALSE)
  
  for (c in 2:10)
    writeFormula(wb, 1, sprintf("%s1 + 1", int2col(c - 1)),
                 startRow = 1, startCol = c)
  
  for (c in 2:10)
    writeFormula(wb, 1, sprintf("%s1 + %s2", int2col(c), int2col(c - 1)),
                 startRow = 2, startCol = c)
  
  for (c in 2:10)
    writeFormula(wb, 1, sprintf("%s2 + %s2", int2col(c), int2col(c + 1)),
                 startRow = 3, startCol = c)
  
  deleteDataColumn(wb, 1, 3)
  # saveWorkbook(wb, "tester.xlsx") # and inspect by hand: expect lots of #REF!
  
  expect_equal(read.xlsx(wb), data.frame(`1` = 1, check.names = FALSE))
  expect_equal(
    wb$worksheets[[1]]$sheet_data$f,
    c(NA, "<f>A1 + 1</f>", "<f>#REF!1 + 1</f>", "<f>C1 + 1</f>", 
      "<f>D1 + 1</f>", "<f>E1 + 1</f>", "<f>F1 + 1</f>", "<f>G1 + 1</f>", 
      "<f>H1 + 1</f>", NA, "<f>B1 + A2</f>", "<f>#REF!1 + #REF!1</f>", 
      "<f>C1 + C1</f>", "<f>D1 + D1</f>", "<f>E1 + E1</f>", "<f>F1 + F1</f>", 
      "<f>G1 + G1</f>", "<f>H1 + H1</f>", "<f>B2 + #REF!1</f>", "<f>C1 + D1</f>", 
      "<f>D1 + E1</f>", "<f>E1 + F1</f>", "<f>F1 + G1</f>", "<f>G1 + H1</f>", 
      "<f>H1 + I1</f>", "<f>I1 + J1</f>")
  )
})