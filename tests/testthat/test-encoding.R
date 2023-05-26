


context("Encoding Tests")



test_that("Write read encoding equality", {
  tempFile <- temp_xlsx()

  wb <- createWorkbook()
  for (i in 1:4) {
    addWorksheet(wb, sprintf("Sheet %s", i))
  }

  df <- data.frame("X" = c("测试", "一下"), stringsAsFactors = FALSE)
  writeDataTable(wb, sheet = 1, x = df)

  saveWorkbook(wb, tempFile, overwrite = TRUE)

  x <- read.xlsx(tempFile)
  expect_equal(x, df)

  x <- read.xlsx(wb)
  expect_equal(x, df)

  ## reload
  wb <- loadWorkbook(tempFile)

  x <- read.xlsx(wb)
  expect_equal(x, df)

  saveWorkbook(wb, tempFile, overwrite = TRUE)
  x <- read.xlsx(tempFile)
  expect_equal(x, df)

  unlink(tempFile, recursive = TRUE, force = TRUE)
  rm(wb)
})


test_that("Support non-ASCII strings not in UTF-8 encodings", {
  non_ascii <- c("\u4f60\u597d", "\u4e2d\u6587", "\u6c49\u5b57", 
                 "\u5de5\u4f5c\u7c3f1", "\u6d4b\u8bd5\u540d\u5b571",
                 "\u6d4b2", "\u5de52", "\u5de5\u4f5c3")
  # Ideally, we should test against native encodings. However, the testing machine's
  #   locale encoding may not be able to represent the non-ascii letters, when
  #   it's the case, we use the UTF-8 encoding as it is.
  if (identical( enc2utf8(enc2native(non_ascii)), non_ascii )) {
    non_ascii <- enc2native(non_ascii)
  }
  non_ascii_df <- data.frame(
    X = non_ascii, Y = seq_along(non_ascii), stringsAsFactors = FALSE
  )
  colnames(non_ascii_df) <- non_ascii[3:4]
  file <- temp_xlsx()
  wb <- createWorkbook(creator = non_ascii[1])
  ws <- addWorksheet(wb, non_ascii[2])
  writeDataTable(wb, ws, non_ascii_df, tableName = non_ascii[3])
  writeData(wb, ws, non_ascii_df, xy = list("D", 1), name = non_ascii[4])
  writeComment(wb, ws, 1, 1, comment = createComment(non_ascii[5], non_ascii[6]))
  writeFormula(wb, ws, x = sprintf('"%s"&"%s"', non_ascii[1], non_ascii[2]), xy = list("G", 1))
  createNamedRegion(wb, ws, 7, 1, name = non_ascii[7])
  saveWorkbook(wb, file)
  
  wb2 <- loadWorkbook(file)
  expect_equal(
    getCreators(wb2), non_ascii[1]
  )
  expect_equal(
    getSheetNames(file), non_ascii[2]
  )
  expect_equivalent(
    getTables(wb2, ws), non_ascii[3]
  )
  expect_equivalent(
    getNamedRegions(wb2), non_ascii[c(4, 7)]
  )
  expect_equal(
    wb2$comments[[1]][[1]][c("comment", "author")],
    setNames(as.list(non_ascii[5:6]), c("comment", "author"))
  )
  expect_equal(
    read.xlsx(file, ws, cols = 1:2),
    non_ascii_df
  )
  expect_equal(
    read.xlsx(file, ws, cols = 4:5),
    non_ascii_df
  )
})
