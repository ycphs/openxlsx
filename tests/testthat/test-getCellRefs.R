

context("Check Cell Ref")



test_that("Provide tests for single getCellRefs", {
  expect_equal(getCellRefs(data.frame(1, 2)), "B1")


  expect_error(getCellRefs(c(1, 2)))

  expect_error(getCellRefs(c(1, "a")))
})


test_that("Provide tests for multiple getCellRefs", {
  expect_equal(getCellRefs(data.frame(1:3, 2:4)), c("B1", "C2", "D3"))




  expect_error(getCellRefs(c(1:2, c("a", "b"))))
})
