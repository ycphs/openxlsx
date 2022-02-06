


context("Group visibility and nested Grouping")



test_that("parsing border xml", {
  wb <- loadWorkbook(file = system.file("extdata", "nested_grouped_rowscols.xlsx", package = "openxlsx"))
  fileName <- file.path(tempdir(), "nested_grouped_rowscols_out.xlsx")
  
  stripAttributes <- function(x) {
    y = x
    attributes(y) <- NULL
    y
  }
  
  ### 1. Loading of outline level
  
  # First tab has simple, visible groupings
  expect_equal(c("1", "1"), stripAttributes(wb$outlineLevels[[1]]))
  expect_equal(c("1", "1"), stripAttributes(wb$colOutlineLevels[[1]]))
  # Second tab has simple, visible groupings
  expect_equal(c("1", "1"), stripAttributes(wb$outlineLevels[[2]]))
  expect_equal(c("1", "1"), stripAttributes(wb$colOutlineLevels[[2]]))
  # Third tab has nested, visible groupings
  expect_equal(c("1", "2", "2", "1"), stripAttributes(wb$outlineLevels[[3]]))
  expect_equal(c("1", "2", "2", "1"), stripAttributes(wb$colOutlineLevels[[3]]))

  
  ### 2. Loading of row/columng visibility of groupings (hidden attribute)
  
  # First tab has visible groupings, second tab has hidden groupings, third has visible, fourth has visible and hidden rows/cols
  expect_equal(rep(NA_character_, 2), attr(wb$outlineLevels[[1]], "hidden"))
  expect_equal(rep("0", 2), attr(wb$colOutlineLevels[[1]], "hidden"))

  expect_equal(c("1", "1"), attr(wb$outlineLevels[[2]], "hidden"))
  expect_equal(c("1", "1"), attr(wb$colOutlineLevels[[2]], "hidden"))
  
  expect_equal(rep(NA_character_, 4), attr(wb$outlineLevels[[3]], "hidden"))
  expect_equal(rep("0", 4), attr(wb$colOutlineLevels[[3]], "hidden"))
  
  expect_equal(c(NA_character_, "1", "1", NA_character_), attr(wb$outlineLevels[[4]], "hidden"))
  expect_equal(c("0", "1", "1", "0"), attr(wb$colOutlineLevels[[4]], "hidden"))
  
  
  #### 3. Export xlsx file and read it in again to check if outlines are preserved
  # The test file has groupings beyond the actual file contents, so not all 
  # grouped rows/cols are available in the data and need to be added manually
  # for the grouping data!
  
  openxlsx::saveWorkbook(wb, file = fileName, overwrite = TRUE)
  wbout <- loadWorkbook(file = fileName)
  
  expect_equal(wb$outlineLevels[[1]], wbout$outlineLevels[[1]]) # BUG: grouping beyond the table's contents are not manually added to the data
  expect_equal(wb$outlineLevels[[2]], wbout$outlineLevels[[2]]) # BUG: grouping beyond the table's contents are not manually added to the data
  expect_equal(wb$outlineLevels[[3]], wbout$outlineLevels[[3]]) # BUG: grouping beyond the table's contents are not manually added to the data
  expect_equal(wb$outlineLevels[[4]], wbout$outlineLevels[[4]])
  
  expect_equal(wb$colOutlineLevels[[1]], wbout$colOutlineLevels[[1]])
  expect_equal(wb$colOutlineLevels[[2]], wbout$colOutlineLevels[[2]])
  expect_equal(wb$colOutlineLevels[[3]], wbout$colOutlineLevels[[3]])
  expect_equal(wb$colOutlineLevels[[4]], wbout$colOutlineLevels[[4]]) # BUG: Ordering of entries not preserved 
  
  
  #### 4. Manually create nested groupings
  
  wb$addWorksheet(sheetName = "Manually")
  writeData(wb, "Manually", matrix(1:225, 25, 25), 1, 1)
  groupRows(wb, "Manually", 2:3, FALSE)
  groupRows(wb, "Manually", 5:6, TRUE)  # BUG: hidden attribute only set for first groupRows call!
  groupRows(wb, "Manually", 9:15) # Outline level 1, next groupings will be level 2! 
  groupRows(wb, "Manually", 10:11)      # BUG: OutlineLevel="2" not properly set (only level "1")
  groupRows(wb, "Manually", 13:14)      # BUG: OutlineLevel="2" not properly set (only level "1")
  expect_equal(c(`2` = "1", `3` = "1", `5` = "1", `6` = "1", `9` = "1", `10` = "2", `11` = "2", `12` = "1", `13` = "2", `14` = "2", `15` = "1"),
               wb$outlineLevels[[5]])

  groupColumns(wb, "Manually", 2:3, FALSE)
  groupColumns(wb, "Manually", 5:6, TRUE)
  groupColumns(wb, "Manually", 9:15)
  groupColumns(wb, "Manually", 10:11)
  groupColumns(wb, "Manually", 13:14)
  expect_equal(c(`2` = "1", `3` = "1", `5` = "1", `6` = "1", `9` = "1", `10` = "2", `11` = "2", `12` = "1", `13` = "2", `14` = "2", `15` = "1"),
               wb$colOutlineLevels[[5]])
  
  
  
  #### CLEANUP
  
  unlink(fileName, recursive = TRUE, force = TRUE)
  
})
