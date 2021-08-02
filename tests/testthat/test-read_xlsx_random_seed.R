test_that("read_xlsx() does not change random seed", {
  rs <- .Random.seed
  expect_identical(rs, .Random.seed)
  tf <- temp_xlsx()
  expect_identical(rs, .Random.seed)
  write.xlsx(data.frame(a = 1), tf)
  expect_identical(rs, .Random.seed)
  read.xlsx(tf)
  expect_identical(rs, .Random.seed)
  unlink(tf)
})
