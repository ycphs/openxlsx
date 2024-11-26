context("silent get_worksheet_entries")

test_that("get_worksheet_entries is silent in reference to string in other sheet", {
  fl <- system.file("extdata", "silent_worksheet_entries.xlsx", package = "openxlsx")
  wb <- loadWorkbook(fl)
  expect_silent(get_worksheet_entries(wb, 1))
})
