context("Workbook groupings")

# groupRows() requires assigning to wb to global environment
# For reference, see here: https://github.com/r-lib/testthat/issues/720

test_that("group columns", {

  # Grouping then setting widths updates hidden
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  groupColumns(wb, "Sheet 1", 2:3, hidden = TRUE)
  setColWidths(wb, "Sheet 1", 2, widths = "18", hidden = FALSE)

  expect_equal(attr(wb$colOutlineLevels[[1]], "hidden")[attr(wb$colOutlineLevels[[1]], "names") == 2], "0")

  # Setting column widths then grouping
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  setColWidths(wb, "Sheet 1", 2:3, widths = "18", hidden = FALSE)
  groupColumns(wb, "Sheet 1", 1:2, hidden = TRUE)

  expect_equal(attr(wb$colWidths[[1]], "hidden")[attr(wb$colWidths[[1]], "names") == 2], "1")
})


test_that("group rows", {

  wb <- createWorkbook()
  assign("wb", wb, envir = .GlobalEnv)
  addWorksheet(wb, "Sheet 1")
  groupRows(wb, "Sheet 1", 1:4, hidden = TRUE)

  expect_equal(names(wb$outlineLevels[[1]]), c("1", "2", "3", "4"))
  expect_equal(unique(attr(wb$outlineLevels[[1]], "hidden")), "1")
  rm(wb)
})

test_that("group rows 2", {
 
  wb <- createWorkbook()
  addWorksheet(wb, 'Sheet1')
  addWorksheet(wb, 'Sheet2')
  writeData(wb, "Sheet1", iris) 
  writeData(wb, "Sheet2", iris)

  ## create list of groups
  # lines used for grouping (here: species)
  grp <- list(
    seq(2, 51),
    seq(52, 101),
    seq(102, 151)
  )
  # assign group levels
  names(grp) <- c("1","0","1") 
  groupRows(wb, "Sheet1", rows = grp)

  # different grouping
  names(grp) <- c("1","2","3")
  groupRows(wb, "Sheet2", rows = grp)

  expect_equal(unique(wb$outlineLevels[[1]]), c("1", "0"))
  expect_equal(unique(wb$outlineLevels[[2]]), c("1", "2", "3"))
})



test_that("ungroup columns", {

  # OutlineLevelCol is removed from SheetFormatPr when no
  # column groupings left
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  setColWidths(wb, "Sheet 1", 2:3, widths = "18", hidden = FALSE)
  groupColumns(wb, "Sheet 1", 1:3, hidden = TRUE)
  ungroupColumns(wb, "Sheet 1", 1:3)

  expect_equal(unique(attr(wb$colWidths[[1]], "hidden")[attr(wb$colWidths[[1]], "names") %in% c(2, 3)]), "0")
})


test_that("ungroup rows", {
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  groupRows(wb, "Sheet 1", 1:3, hidden = TRUE)
  ungroupRows(wb, "Sheet 1", 1:3)

  expect_equal(length(wb$outlineLevels[[1]]), 0L)
})

test_that("no warnings #485", {

  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")

  expect_silent(
    groupRows(
      rows = c(1),
      sheet = "Sheet 1",
      wb = wb
    )
  )

  expect_silent(
    ungroupRows(
      rows = c(1),
      sheet = "Sheet 1",
      wb = wb
    )
  )

})


test_that("loading workbook preserves outlines", {
  fl <- system.file("extdata", "groupTest.xlsx", package = "openxlsx")
  wb <- loadWorkbook(fl)

  expect_equal(names(wb$colOutlineLevels[[1]]), c("2", "3", "4"))
  expect_equal(names(wb$outlineLevels[[1]]), c("3", "4"))
  expect_equal(unique(attr(wb$outlineLevels[[1]], "hidden")[names(wb$outlineLevels[[1]]) %in% c("3", "4")]), "true")

  wbb <- createWorkbook()

  addWorksheet(wbb, sheetName = "Test", gridLines = FALSE, tabColour = "deepskyblue")

  writeData(wbb, sheet = "Test", x = c(colA = "testcol1", colB = "testcol2"))

  groupColumns(wbb, "Test", cols = 2:3, hidden = FALSE)
  setColWidths(wbb, sheet = "Test", cols=c(1:5), widths = c(9,9,9,9,9))
  groupColumns(wbb, "Test", cols = 5:10, hidden = FALSE)
  setColWidths(wbb, "Test", cols = 15:20, widths = 9)

  tf <- temp_xlsx("test")
  tf2 <- temp_xlsx("test2")

  saveWorkbook(wbb, tf, overwrite = TRUE)
  test <- wbb$worksheets[[1]]$copy()

  wb <- loadWorkbook(tf)
  saveWorkbook(wb, tf2, overwrite = TRUE)

  testthat::expect_equal(wb$worksheets[[1]], test)

  unlink(c("tf", "tf2"), recursive = TRUE, force = TRUE)
})


test_that("Grouping after setting colwidths has correct length of hidden attributes", {
  # Issue #100 - https://github.com/ycphs/openxlsx/issues/100

  wb <- createWorkbook(title = "column width and grouping error")
  addWorksheet(wb, sheetName = 1)

  setColWidths(
    wb,
    sheet = 1,
    cols = 1:100,
    widths = 8
  )

  groupColumns(wb, sheet = 1, cols = 20:100, hidden = TRUE)

  expect_equal(length(wb$colOutlineLevels[[1]]), length(attr(wb$colOutlineLevels[[1]], "hidden")))
})

test_that("Consecutive calls to saveWorkbook doesn't corrupt attributes", {

  wbb <- createWorkbook()

  addWorksheet(wbb, sheetName = "Test", gridLines = FALSE, tabColour = "deepskyblue")

  writeData(wbb, sheet = "Test", x = c(colA = "testcol1", colB = "testcol2"))

  groupColumns(wbb, "Test", cols = 2:3, hidden = FALSE)
  setColWidths(wbb, sheet = "Test", cols=c(1:5), widths = c(9,9,9,9,9))
  groupColumns(wbb, "Test", cols = 5:10, hidden = FALSE)
  setColWidths(wbb, "Test", cols = 15:20, widths = 9)

  tf <- temp_xlsx("test")
  tf2 <- temp_xlsx("test2")

  saveWorkbook(wbb, tf, overwrite = TRUE)
  test <- wbb$worksheets[[1]]$copy()

  saveWorkbook(wbb, tf2, overwrite = TRUE)

  testthat::expect_equal(wbb$worksheets[[1]], test)

  unlink(c("tf", "tf2"), recursive = TRUE, force = TRUE)
})
