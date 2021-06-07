test_that("lookup tables dimensions", {
  expect_equal(dim(openxlsxFontSizeLookupTable), c(29L, 225L))
  expect_equal(dim(openxlsxFontSizeLookupTableBold), c(29L, 225L))
})
