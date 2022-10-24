


context("Group visibility and nested Grouping")



test_that("row and column grouping", {
  # Test file has several worksheets:
  #   1. "simple visible grouping": One simple grouping (rows 1:2, columns 1:2), visible
  #   2. "simple hidden grouping": One simple grouping (rows 1:2, columns 1:2), hidden
  #   3. "multiple non-nested groupings": three groupings (rows/cols 1:2, 4:5 and 7:8 (hidden); table contents only until cell G6, so last grouping has no actual content)
  #   4. "nested grouping": cols rows 1:4 (outline level 1) and 2:3 (outline level 2), contents to cell F7 (spanning all groupings)
  #   4. "nested grouping 2": Like previous, but outline level 2 is hidden, contents only to cell B2
  
  # Potential issues with groupings:
  #   - Outline level not preserved
  #   - Hidden flag not set (or not correctly loaded)
  #   - When nesting goes beyond the cell contents, rows/cols need to be added manually to the xml for the grouping settings
  #   - Hidden flag for visible rows/cols not properly loaded, so subsequent hidden cols set the hidden flag for the wrong row/col
  #   - Non-overlapping groupings 

  
  wb <- loadWorkbook(file = system.file("extdata", "nested_grouped_rowscols.xlsx", package = "openxlsx"))
  fileName <- file.path(tempdir(), "nested_grouped_rowscols_out.xlsx")
  oL1 = c(`1` = "1", `2` = "1")
  oL3 = c(`1` = "1", `2` = "1", `4` = "1", `5` = "1", `7` = "1", `8` = "1")
  oL4 = c(`1` = "1", `2` = "2", `3` = "2", `4` = "1")
  
  stripAttributes <- function(x, attr = NA) {
    y = x
    if (is.na(attr)) {
      attributes(y) <- NULL
    } else {
      attr(y, attr) <- NULL
    }
    y
  }
  
  ### 1. Loading of outline level
  
  # First sheet has simple, visible groupings
  expect_equal(oL1, stripAttributes(wb$outlineLevels[[1]], "hidden"))
  expect_equal(oL1, stripAttributes(wb$colOutlineLevels[[1]], "hidden"))
  # Second sheet has simple, hidden groupings
  expect_equal(oL1, stripAttributes(wb$outlineLevels[[2]], "hidden"))
  expect_equal(oL1, stripAttributes(wb$colOutlineLevels[[2]], "hidden"))
  # Third sheet has simple, non-overlapping groupings
  expect_equal(oL3, stripAttributes(wb$outlineLevels[[3]], "hidden"))
  expect_equal(oL3, stripAttributes(wb$colOutlineLevels[[3]], "hidden"))
  # Fourth tab has nested, visible groupings
  expect_equal(oL4, stripAttributes(wb$outlineLevels[[4]], "hidden"))
  expect_equal(oL4, stripAttributes(wb$colOutlineLevels[[4]], "hidden"))

  
  ### 2. Loading of row/columng visibility of groupings (hidden attribute)
  
  # First tab has visible groupings, second tab has hidden groupings, third has visible, fourth has visible and hidden rows/cols
  expect_equal(rep(NA_character_, 2), attr(wb$outlineLevels[[1]], "hidden"))
  expect_equal(c("1", "1"), attr(wb$outlineLevels[[2]], "hidden"))
  expect_equal(c(rep(NA_character_, 4), "1", "1"), attr(wb$outlineLevels[[3]], "hidden"))
  expect_equal(rep(NA_character_, 4), attr(wb$outlineLevels[[4]], "hidden"))
  expect_equal(c(NA_character_, "1", "1", NA_character_), attr(wb$outlineLevels[[5]], "hidden"))
  
  expect_equal(rep("0", 2), attr(wb$colOutlineLevels[[1]], "hidden"))
  expect_equal(c("1", "1"), attr(wb$colOutlineLevels[[2]], "hidden"))
  expect_equal(c("0", "0", "0", "0", "1", "1"), attr(wb$colOutlineLevels[[3]], "hidden"))
  expect_equal(rep("0", 4), attr(wb$colOutlineLevels[[4]], "hidden"))
  expect_equal(c("0", "1", "1", "0"), attr(wb$colOutlineLevels[[5]], "hidden"))
  
  
  #### 3. Export xlsx file and read it in again to check if outlines are preserved
  # The test file has groupings beyond the actual file contents, so not all 
  # grouped rows/cols are available in the data and need to be added manually
  # for the grouping data!
  
  openxlsx::saveWorkbook(wb, file = fileName, overwrite = TRUE)
  wbout <- loadWorkbook(file = fileName)
  
  expect_equal(wb$outlineLevels[[1]], wbout$outlineLevels[[1]])
  expect_equal(wb$outlineLevels[[2]], wbout$outlineLevels[[2]])
  expect_equal(wb$outlineLevels[[3]], wbout$outlineLevels[[3]])
  expect_equal(wb$outlineLevels[[4]], wbout$outlineLevels[[4]])
  expect_equal(wb$outlineLevels[[5]], wbout$outlineLevels[[5]])
  
  expect_equal(wb$colOutlineLevels[[1]], wbout$colOutlineLevels[[1]])
  expect_equal(wb$colOutlineLevels[[2]], wbout$colOutlineLevels[[2]])
  expect_equal(wb$colOutlineLevels[[3]], wbout$colOutlineLevels[[3]]) # BUG: Ordering of entries not preserved
  expect_equal(wb$colOutlineLevels[[4]], wbout$colOutlineLevels[[4]]) # BUG: Ordering of entries not preserved
  expect_equal(wb$colOutlineLevels[[5]], wbout$colOutlineLevels[[5]]) # BUG: Ordering of entries not preserved
  
  
  #### 4. Manually create non-overlapping and nested groupings of rows / cols
  
  wb <- openxlsx::createWorkbook()
  wb$addWorksheet(sheetName = "Manually")
  writeData(wb, "Manually", matrix(1:225, 15, 15), 1, 1)
  # non-nested grouping (visible and hidden)
  groupRows(wb, "Manually", 2:3, FALSE)
  groupRows(wb, "Manually", 5:6, TRUE)
  # nested grouping
  groupRows(wb, "Manually", 9:15) # Outline level 1, next groupings will be level 2! 
  groupRows(wb, "Manually", 10:11)
  groupRows(wb, "Manually", 13:14)
  expect_equal(c(`2` = "1", `3` = "1", `5` = "1", `6` = "1", `9` = "1", `10` = "2", `11` = "2", `12` = "1", `13` = "2", `14` = "2", `15` = "1"),
               stripAttributes(wb$outlineLevels[[1]], "hidden"))

  # non-nested grouping (visible and hidden)
  groupColumns(wb, "Manually", 2:3, FALSE)
  groupColumns(wb, "Manually", 5:6, TRUE)
  # nested grouping
  groupColumns(wb, "Manually", 9:15)
  groupColumns(wb, "Manually", 10:11)
  groupColumns(wb, "Manually", 13:14)
  expect_equal(c(`2` = "1", `3` = "1", `5` = "1", `6` = "1", `9` = "1", `10` = "2", `11` = "2", `12` = "1", `13` = "2", `14` = "2", `15` = "1"),
               stripAttributes(wb$colOutlineLevels[[1]], "hidden"))
  
  
  #### 5. Ungrouping rows/cols simply decrements the level (and removes the entry if no longer grouped)
  ungroupRows(wb, "Manually", 11:13)
  expect_equal(c(`2` = "1", `3` = "1", `5` = "1", `6` = "1", `9` = "1", `10` = "2", `11` = "1", `13` = "1", `14` = "2", `15` = "1"),
               stripAttributes(wb$outlineLevels[[1]], "hidden"))
  
  ungroupColumns(wb, "Manually", 11:13)
  expect_equal(c(`2` = "1", `3` = "1", `5` = "1", `6` = "1", `9` = "1", `10` = "2", `11` = "1", `13` = "1", `14` = "2", `15` = "1"),
               stripAttributes(wb$colOutlineLevels[[1]], "hidden"))
  

  #### CLEANUP
  
  unlink(fileName, recursive = TRUE, force = TRUE)
  
})
