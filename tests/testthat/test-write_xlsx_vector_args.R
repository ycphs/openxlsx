
context("write.xlsx vector arguments")

test_that("Writing then reading returns identical data.frame 1", {
  tmp_file <- file.path(tempdir(), "xlsx_vector_args.xlsx")

  df1 <- data.frame(1:2)
  df2 <- data.frame(1:3)
  x <- list(df1, df2)

  write.xlsx(
    file = tmp_file,
    x = x,
    gridLines = c(F, T),
    sheetName = c("a", "b"),
    zoom = c(50, 90),
    tabColour = c("red", "blue")
  )

  wb <- loadWorkbook(tmp_file)

  expect_equal(object = getSheetNames(tmp_file), expected = c("a", "b"))
  expect_equal(object = names(wb), expected = c("a", "b"))

  expect_true(object = grepl('rgb="FFFF0000"', wb$worksheets[[1]]$sheetPr))
  expect_true(object = grepl('rgb="FF0000FF"', wb$worksheets[[2]]$sheetPr))

  expect_true(object = grepl('zoomScale="50"', wb$worksheets[[1]]$sheetViews))
  expect_true(object = grepl('zoomScale="90"', wb$worksheets[[2]]$sheetViews))

  expect_true(object = grepl('showGridLines="0"', wb$worksheets[[1]]$sheetViews))
  expect_true(object = grepl('showGridLines="1"', wb$worksheets[[2]]$sheetViews))

  expect_equal(read.xlsx(tmp_file, sheet = 1), df1)
  expect_equal(read.xlsx(tmp_file, sheet = 2), df2)

  unlink(tmp_file, recursive = TRUE, force = TRUE)
})

test_that("write.xlsx() passes withFilter and colWidths [151]", {
  df <- data.frame(x = 1, b = 2)
  tf1 <- tempfile("file_1_", fileext = ".xlsx")
  tf2 <- tempfile("file_2_", fileext = ".xlsx")
  on.exit(file.remove(tf1, tf2), add = TRUE)
  
  # undebug(write.xlsx)
  # withFilter default should be FALSE when asTable is FALSE
  write.xlsx(df, tf1) 
  write.xlsx(df, tf2, withFilter = TRUE, colWidths = 15)
  
  x <- loadWorkbook(tf1)
  y <- loadWorkbook(tf2)
  
  expect_equal(
    x$worksheets[[1]]$autoFilter,
    character()
  )
  expect_equal(
    y$worksheets[[1]]$autoFilter,
    "<autoFilter ref=\"A1:B2\"/>"
  )
  
  expect_equal(x$colWidths, list(list()))
  
  expect_equal(
    y$colWidths[[1]],
    structure(c(`1` = "15", `2` = "15"), hidden = c("0", "0"))
  )
}) 

test_that("write.xlsx() correctly passes default asTable and withFilters", {
  df <- data.frame(x = 1, b = 2)
  tf1 <- tempfile(fileext = ".xlsx")
  tf2 <- tempfile(fileext = ".xlsx")
  on.exit(file.remove(tf1, tf2), add = TRUE)
  
  # asTable = TRUE  >> writeDataTable >> withFilter = TRUE
  # asTable = FALSE >> writeData      >> withFilter = FALSE
  write.xlsx(df, tf1, asTable = FALSE) 
  write.xlsx(df, tf2, asTable = TRUE)
  
  x <- loadWorkbook(tf1)
  y <- loadWorkbook(tf2)
  
  expect_equal(
    x$worksheets[[1]]$autoFilter,
    character()
  )
  
  # not autoFilter for tables
  expect_equal(
    y$worksheets[[1]]$tableParts,
    structure("<tablePart r:id=\"rId3\"/>", tableName = c(`A1:B2` = "Table3"))
  )
})

test_that("write.xlsx() correctly handles colWidths", {
  x <- data.frame(a = 1, b = 2, c = 3)
  file <- tempfile("write_xlsx_", fileext = ".xlsx")
  on.exit(if (file.exists(file)) file.remove(file), add = TRUE)
  zero3 <- rep("0", 3)
  
  # No warning when passing "auto"
  expect_warning(write.xlsx(rep_len(list(x), 3), file, colWidths = "auto"), NA)
  
  # single value is repeated for all columns
  write.xlsx(rep_len(list(x), 3), file, colWidths = 13)
  
  expect_equal(
    loadWorkbook(file)$colWidths,
    rep_len(list(structure(c(`1` = "13", `2` = "13", `3` = "13"), hidden = zero3)), 3)
  )
  
  # sets are repated
  write.xlsx(rep_len(list(x), 3), file, colWidths = list(c(10, 20, 30)))
  
  expect_equal(
    loadWorkbook(file)$colWidths,
    rep_len(list(structure(c(`1` = "10", `2` = "20", `3` = "30"), hidden = zero3)), 3)
  )
  
  # 3 distinct sets
  write.xlsx(rep_len(list(x), 3), file, 
    colWidths = list(
      c(10, 20, 30),
      c(100, 200, 300),
      c(1, 2, 3)
    ))
  
  expect_equal(
    loadWorkbook(file)$colWidths,
    list(
      structure(c(`1` = "10", `2` = "20", `3` = "30"), hidden = zero3),
      structure(c(`1` = "100", `2` = "200", `3` = "300"), hidden = zero3),
      structure(c(`1` = "1", `2` = "2", `3` = "3"), hidden = zero3)
    )
  )
})
