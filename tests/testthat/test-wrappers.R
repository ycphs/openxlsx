
context("Test wrappers")

test_that("int2col and col2int", {
  
  nums <- 2:27
  
  chrs <- c("B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",  "N",
            "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA")
  
  expect_equal(chrs, int2col(nums))
  expect_equal(nums, col2int(chrs))
  
})
