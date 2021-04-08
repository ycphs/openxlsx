test_that("option names are appropriate", {
  bad <- grep("^openxlsx[.].*", names(op.openxlsx), value = TRUE, invert = TRUE)
  expect_equal(bad, character(0))
})

test_that("changing options", {
  op <- options()
  on.exit(options(op), add = TRUE)
  
  # Set via options()
  options(openxlsx.border = "whatever")
  expect_equal(openxlsx_getOp("borders"), getOption("openxlsx.borders"))
  
  options(openxlsx.tableStyle = "Cool format")
  expect_equal(openxlsx_getOp("tableStyle"), openxlsx_getOp("openxlsx.tableStyle"))
  
  # Setting to NULL will return default
  options(openxlsx.border = NULL)
  expect_equal(openxlsx_getOp("borders"), op.openxlsx[["openxlsx.borders"]])
  
  # Bad options names will trigger warning but still be produced
  options(openxlsx.likelyNotARealOption = TRUE)
  expect_warning(
    expect_true(openxlsx_getOp("likelyNotARealOption")),
    "not a standard openxlsx option"
  )
  
  # Must be single value
  expect_error(openxlsx_getOp(NULL), "length")
  expect_error(openxlsx_getOp(c("border", "withFilter")), "length")
})
