
test_that("option names are appropriate", {
  bad <- grep("^openxlsx[.].*", names(op.openxlsx), value = TRUE, invert = TRUE)
  expect_equal(bad, character(0))
})

test_that("changing options", {
  op <- options()
  
  # Set via options()
  options(openxlsx.borders = "whatever")
  expect_equal(openxlsx_getOp("borders"), getOption("openxlsx.borders"))
  expect_equal(openxlsx_getOp("borders"), "whatever")
  
  # Set via openxlsx_setOp()
  openxlsx_setOp("borders", "new_whatever")
  expect_equal(openxlsx_getOp("borders"), getOption("openxlsx.borders"))
  expect_equal(openxlsx_getOp("borders"), "new_whatever")
  
  # with openxlsx. prefix
  openxlsx_setOp("openxlsx.borders", "new_new_whatever")
  expect_equal(openxlsx_getOp("openxlsx.borders"), getOption("openxlsx.borders"))
  expect_equal(openxlsx_getOp("openxlsx.borders"), "new_new_whatever")
  
  options(openxlsx.tableStyle = "Cool format")
  expect_equal(openxlsx_getOp("tableStyle"), openxlsx_getOp("openxlsx.tableStyle"))
  
  # Setting to NULL will return default
  options(openxlsx.borders = NULL)
  expect_equal(openxlsx_getOp("borders"), op.openxlsx[["openxlsx.borders"]])
  
  # Bad options names will trigger warning but still be produced
  options(openxlsx.likelyNotARealOption = TRUE)
  expect_warning(
    expect_true(openxlsx_getOp("likelyNotARealOption")),
    "not a standard openxlsx option"
  )
  
  # Multiple Ops returns error
  expect_error(openxlsx_getOp(c("withFilter", "borders")), "length 1")
  
  openxlsx_resetOp()
  options(op)
})

test_that("openxlsx_setOp() works with list [#215]", {
  op <- options()
  expect_error(openxlsx_setOp(list(withFilter = TRUE, keepNA = TRUE)), NA)
  openxlsx_resetOp()
  options(op)
})
