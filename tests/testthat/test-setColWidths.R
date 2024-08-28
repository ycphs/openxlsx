
context("Setting column widths")

test_that("Resetting col widths", {
  # Write some data to a workbook object.
  wb <- createWorkbook()
  addWorksheet(wb, "iris") 
  writeData(wb, "iris", iris)
  
  # Set column widths and perform the pre-save operation to prepare col xml 
  # (typically called by `saveWorkbook()`).
  setColWidths(wb, "iris", cols = 1:2, 12)
  wb$setColWidths(1)
  
  # Set column widths again for a different range inclusive of the previous one, 
  # perform the pre-save operation to prepare col xml, and test for an error 
  # (reported in https://github.com/ycphs/openxlsx/issues/493).
  setColWidths(wb, "iris", cols = 1:5, 15)
  expect_error(wb$setColWidths(1), NA)
  
  # Test that the resulting col xml meets expectations.
  expected_col_xml <- c(
    "1" = "<col min=\"1\" max=\"1\" width=\"15.71\" hidden=\"0\" customWidth=\"1\"/>",
    "2" = "<col min=\"2\" max=\"2\" width=\"15.71\" hidden=\"0\" customWidth=\"1\"/>",
    "3" = "<col min=\"3\" max=\"3\" width=\"15.71\" hidden=\"0\" customWidth=\"1\"/>",
    "4" = "<col min=\"4\" max=\"4\" width=\"15.71\" hidden=\"0\" customWidth=\"1\"/>",
    "5" = "<col min=\"5\" max=\"5\" width=\"15.71\" hidden=\"0\" customWidth=\"1\"/>"
  )
  expect_equal(
    wb$worksheets[[1]]$cols, 
    expected_col_xml
  )
})
