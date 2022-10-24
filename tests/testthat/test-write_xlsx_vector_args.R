
context("write.xlsx vector arguments")

test_that("Writing then reading returns identical data.frame 1", {
  tmp_file <- file.path(tempdir(), "xlsx_vector_args.xlsx")
  
  df1 <- data.frame(1:2)
  df2 <- data.frame(1:3)
  x <- list(df1, df2)
  
  write.xlsx(
    file = tmp_file,
    x = x,
    gridLines = c(FALSE, TRUE),
    sheetName = c("a", "b"),
    zoom = c(50, 90),
    tabColour = c("red", "blue")
  )
  
  wb <- loadWorkbook(tmp_file)
  
  expect_equal(getSheetNames(tmp_file), expected = c("a", "b"))
  expect_equal(names(wb), expected = c("a", "b"))
  
  expect_true(grepl('rgb="FFFF0000"', wb$worksheets[[1]]$sheetPr))
  expect_true(grepl('rgb="FF0000FF"', wb$worksheets[[2]]$sheetPr))
  
  expect_true(grepl('zoomScale="50"', wb$worksheets[[1]]$sheetViews))
  expect_true(grepl('zoomScale="90"', wb$worksheets[[2]]$sheetViews))
  
  expect_true(grepl('showGridLines="0"', wb$worksheets[[1]]$sheetViews))
  expect_true(grepl('showGridLines="1"', wb$worksheets[[2]]$sheetViews))
  
  expect_equal(read.xlsx(tmp_file, sheet = 1), df1)
  expect_equal(read.xlsx(tmp_file, sheet = 2), df2)
  
  unlink(tmp_file, recursive = TRUE, force = TRUE)
})

test_that("write.xlsx() passes withFilter and colWidths [151]", {
  df <- data.frame(x = 1, b = 2)
  x <- buildWorkbook(df) 
  y <- buildWorkbook(df, withFilter = TRUE, colWidths = 15)
  
  expect_equal(x$worksheets[[1]]$autoFilter, character())
  expect_equal(y$worksheets[[1]]$autoFilter, "<autoFilter ref=\"A1:B2\"/>")
  expect_equal(x$colWidths, list(list()))
  
  expect_equal(
    y$colWidths[[1]],
    structure(c(`1` = "15", `2` = "15"), hidden = c("0", "0"))
  )
}) 

test_that("write.xlsx() correctly passes default asTable and withFilters", {
  df <- data.frame(x = 1, b = 2)
  
  # asTable = TRUE  >> writeDataTable >> withFilter = TRUE
  # asTable = FALSE >> writeData      >> withFilter = FALSE
  x <- buildWorkbook(df, asTable = FALSE) 
  y <- buildWorkbook(df, asTable = TRUE)
  
  # Save the workbook 
  tf <- temp_xlsx()
  saveWorkbook(y, tf)
  y2 <- loadWorkbook(tf)
  
  expect_identical(x$worksheets[[1]]$autoFilter, character())
  
  # not autoFilter for tables -- not named in buildWorkbook
  expect_equal(
    y$worksheets[[1]]$tableParts,
    structure("<tablePart r:id=\"rId3\"/>", tableName = "Table3")
  )
  
  expect_equal(
    y2$worksheets[[1]]$tableParts,
    structure("<tablePart r:id=\"rId3\"/>", tableName = c(`A1:B2` = "Table3"))
  )
  
  file.remove(tf)
})

test_that("write.xlsx() correctly handles colWidths", {
  x <- data.frame(a = 1, b = 2, c = 3)
  zero3 <- rep("0", 3)
  
  # No warning when passing "auto"
  expect_warning(buildWorkbook(rep_len(list(x), 3), colWidths = "auto"), NA)
  
  # single value is repeated for all columns
  wb <- buildWorkbook(rep_len(list(x), 3), colWidths = 13)
  exp <- rep_len(list(structure(c(`1` = "13", `2` = "13", `3` = "13"), hidden = zero3)), 3)
  expect_equal(wb$colWidths, exp)
  
  # sets are repeated
  wb <- buildWorkbook(rep_len(list(x), 3), colWidths = list(c(10, 20, 30)))
  exp <- rep_len(list(structure(c(`1` = "10", `2` = "20", `3` = "30"), hidden = zero3)), 3)
  expect_equal(wb$colWidths, exp)
  
  # 3 distinct sets
  wb <- buildWorkbook(
    rep_len(list(x), 3),
    colWidths = list(
      c(10, 20, 30),
      c(100, 200, 300),
      c(1, 2, 3)
    ))
  
  expect_equal(
    wb$colWidths,
    list(
      structure(c(`1` = "10", `2` = "20", `3` = "30"), hidden = zero3),
      structure(c(`1` = "100", `2` = "200", `3` = "300"), hidden = zero3),
      structure(c(`1` = "1", `2` = "2", `3` = "3"), hidden = zero3)
    )
  )
})
