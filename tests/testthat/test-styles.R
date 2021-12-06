
context("Styles")

test_that("setStyle", {
  
  tmp_file <- temp_xlsx()
  
  # lorem ipsum
  txt <- paste0(
    "Lorem ipsum dolor sit amet, consectetur adipiscing elit, ",
    "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ",
    "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris",
    "nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in ",
    "reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla",
    "pariatur. Excepteur sint occaecat cupidatat non proident, sunt in ",
    "culpa qui officia deserunt mollit anim id est laborum."
  )
  
  ## create workbook
  wb <- createWorkbook()
  addWorksheet(wb, "Test")
  writeData(wb, "Test", txt)
  
  ## create a style
  s <- createStyle(
    fontSize = 12, 
    fontColour = "black", 
    valign="center", 
    wrapText = TRUE, 
    halign = "justify"
  )
  addStyle(wb, "Test", s, 1, 1)
  setColWidths(wb, "Test", 1, 50)
  setRowHeights(wb, "Test", 1, 150)
  
  ## save workbook
  saveWorkbook(wb, tmp_file)
  
  ## load it again
  wb2 <- loadWorkbook(tmp_file)
  s2 <- getStyles(wb2)[[1]]
  
  ## test that the style survived the round trip
  expect_equal(s2$fontSize, c(val="12"))
  expect_equal(s2$fontColour, c(rgb="FF000000"))
  expect_equal(s2$valign, "center")
  expect_equal(s2$wrapText, TRUE)
  expect_equal(s2$halign, "justify")
  
})
