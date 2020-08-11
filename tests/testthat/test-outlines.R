context("Workbook groupings")

# groupRows() requires assigning to wb to global environment
# For reference, see here: https://github.com/r-lib/testthat/issues/720

test_that("group columns", {

  # Grouping then setting widths updates hidden
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  groupColumns(wb, "Sheet 1", 2:3, hidden = T)
  setColWidths(wb, "Sheet 1", 2, widths = "18", hidden = F)

  expect_equal(attr(wb$colOutlineLevels[[1]], "hidden")[attr(wb$colOutlineLevels[[1]], "names") == 2], "0")

  # Setting column widths then grouping
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  setColWidths(wb, "Sheet 1", 2:3, widths = "18", hidden = F)
  groupColumns(wb, "Sheet 1", 1:2, hidden = T)

  expect_equal(attr(wb$colWidths[[1]], "hidden")[attr(wb$colWidths[[1]], "names") == 2], "1")
})


test_that("group rows", {
  wb <- createWorkbook()
  assign("wb", wb, envir = .GlobalEnv)
  addWorksheet(wb, "Sheet 1")
  groupRows(wb, "Sheet 1", 1:4, hidden = T)

  expect_equal(names(wb$outlineLevels[[1]]), c("1", "2", "3", "4"))
  expect_equal(unique(attr(wb$outlineLevels[[1]], "hidden")), "1")
  rm(wb)
})


test_that("ungroup columns", {

  # OutlineLevelCol is removed from SheetFormatPr when no
  # column groupings left
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  setColWidths(wb, "Sheet 1", 2:3, widths = "18", hidden = F)
  groupColumns(wb, "Sheet 1", 1:3, hidden = T)
  ungroupColumns(wb, "Sheet 1", 1:3)

  expect_equal(unique(attr(wb$colWidths[[1]], "hidden")[attr(wb$colWidths[[1]], "names") %in% c(2, 3)]), "0")
})


test_that("ungroup rows", {
  wb <- createWorkbook()
  assign("wb", wb, envir = .GlobalEnv)
  addWorksheet(wb, "Sheet 1")
  groupRows(wb, "Sheet 1", 1:3, hidden = T)
  ungroupRows(wb, "Sheet 1", 1:3)

  expect_equal(length(wb$outlineLevels[[1]]), 0L)
  rm(wb)
})


test_that("loading workbook preserves outlines", {
  fl <- system.file("extdata", "groupTest.xlsx", package = "openxlsx")
  wb <- loadWorkbook(fl)

  expect_equal(names(wb$colOutlineLevels[[1]]), c("2", "3", "4"))
  expect_equal(names(wb$outlineLevels[[1]]), c("3", "4"))
  expect_equal(unique(attr(wb$outlineLevels[[1]], "hidden")[names(wb$outlineLevels[[1]]) %in% c("3", "4")]), "true")
})
