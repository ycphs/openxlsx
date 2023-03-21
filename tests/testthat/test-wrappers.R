
context("Test wrappers")

test_that("int2col and col2int", {
  
  nums <- 2:27
  
  chrs <- c("B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",  "N",
            "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA")
  
  expect_equal(chrs, int2col(nums))
  expect_equal(nums, col2int(chrs))
  
})


test_that("deleteDataColumn basics", {
  wb <- createWorkbook()
  addWorksheet(wb, "tester")

  for (i in seq(5)) {
    mat <- data.frame(x = rep(paste0(int2col(i), i), 10))
    writeData(wb, sheet = 1, startRow = 1, startCol = i, mat)
    writeFormula(wb, sheet = 1, startRow = 12, startCol = i,
                 x = sprintf("=COUNTA(%s2:%s11)", int2col(i), int2col(i)))
  }
  expect_equal(
    setdiff(wb$worksheets[[1]]$sheet_data$f, NA),
    c("<f>=COUNTA(A2:A11)</f>", "<f>=COUNTA(B2:B11)</f>", "<f>=COUNTA(C2:C11)</f>", 
      "<f>=COUNTA(D2:D11)</f>", "<f>=COUNTA(E2:E11)</f>")
  )


  deleteDataColumn(wb, 1, col = 3)
  expect_equal(read.xlsx(wb),
               data.frame(x = rep("A1", 10), x = "B2", x = "D4", x = "E5", # no C3!
                          check.names = FALSE))
  expect_equal(
    setdiff(wb$worksheets[[1]]$sheet_data$f, NA),
    c("<f>=COUNTA(A2:A11)</f>", "<f>=COUNTA(B2:B11)</f>", 
      "<f>=COUNTA(C2:C11)</f>", "<f>=COUNTA(D2:D11)</f>")
  )
  
  
  deleteDataColumn(wb, 1, col = 2)
  expect_equal(read.xlsx(wb),
               data.frame(x = rep("A1", 10), x = "D4", x = "E5", # no B2!
                          check.names = FALSE))
  expect_equal(
    setdiff(wb$worksheets[[1]]$sheet_data$f, NA),
    c("<f>=COUNTA(A2:A11)</f>", "<f>=COUNTA(B2:B11)</f>", "<f>=COUNTA(C2:C11)</f>")
  )
  
  deleteDataColumn(wb, 1, col = 1)
  expect_equal(read.xlsx(wb),
               data.frame(x = rep("D4", 10), x = "E5", # no A1!
                          check.names = FALSE))
  expect_equal(
    setdiff(wb$worksheets[[1]]$sheet_data$f, NA),
    c("<f>=COUNTA(A2:A11)</f>", "<f>=COUNTA(B2:B11)</f>")
  )
})


test_that("deleteDataColumn with more complicated formulae", {
  # works with more complicated formula as well!
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
    c(NA, "<f>A1 + 1</f>", "<f>#REF!1 + 1</f>", "<f>C1 + 1</f>", "<f>D1 + 1</f>", "<f>E1 + 1</f>", "<f>F1 + 1</f>", "<f>G1 + 1</f>", "<f>H1 + 1</f>",
      NA, "<f>B1 + A2</f>", "<f>C1 + #REF!2</f>", "<f>D1 + C2</f>", "<f>E1 + D2</f>", "<f>F1 + E2</f>", "<f>G1 + F2</f>", "<f>H1 + G2</f>", "<f>I1 + H2</f>",
      "<f>B2 + #REF!2</f>", "<f>C2 + D2</f>", "<f>D2 + E2</f>", "<f>E2 + F2</f>", "<f>F2 + G2</f>", "<f>G2 + H2</f>", "<f>H2 + I2</f>", "<f>I2 + J2</f>")
  )
})


test_that("deleteDataColumn with wide data", {
  wb <- createWorkbook()
  addWorksheet(wb, "tester")
  ncols <- 30
  nrows <- 100
  df <- as.data.frame(matrix(seq(ncols * nrows), ncol = ncols))
  colnames(df) <- int2col(seq(ncols))
  writeData(wb, sheet = 1, startRow = 1, startCol = 1, x = df, colNames = TRUE)
  expect_equal(read.xlsx(wb), df)
  
  deleteDataColumn(wb, 1, 2)
  expect_equal(read.xlsx(wb), df[, -2])
  
  deleteDataColumn(wb, 1, 100)
  expect_equal(read.xlsx(wb), df[, -2])
  
  deleteDataColumn(wb, 1, 55)
  expect_equal(read.xlsx(wb), df[, -c(2, 56)]) # 56 b.c. one col was already taken out
  
  deleteDataColumn(wb, 1, 1)
  expect_equal(read.xlsx(wb), df[, -c(1, 2, 56)])
  
  # delete all data
  for (i in seq(ncols - 2))
    deleteDataColumn(wb, 1, 1)
  expect_warning(read.xlsx(wb), "No data found on worksheet")
})

test_that("deleteDataColumn with formatting data", {
  wb <- createWorkbook()
  addWorksheet(wb, "tester")
  df <- data.frame(x = 1:10, y = letters[1:10], z = 10:1)
  writeData(wb, sheet = 1, startRow = 1, startCol = 1, x = df, colNames = TRUE)
  
  st <- openxlsx::createStyle(textDecoration = "Bold", fontSize = 20, fontColour = "red")
  openxlsx::addStyle(wb, 1, style = st, rows = 1, cols = seq(ncol(df)))

  sst <- wb$styleObjects[[1]]
  sst$rows <- c(1, 1)
  sst$cols <- c(1, 2)
  deleteDataColumn(wb, 1, 2)
  
  expect_length(wb$styleObjects, 1)
  expect_equal(wb$styleObjects[[1]],
               sst)
})
