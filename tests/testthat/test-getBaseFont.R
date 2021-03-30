test_that("getBaseFont works", {
  wb <- createWorkbook()
  expect_equal(
    getBaseFont(wb),
    list(
      size = list(val = "11"),
      # should this be "#000000"?
      colour = list(rgb = "FF000000"),
      name = list(val = "Calibri")
    )
  )
  
  modifyBaseFont(wb, fontSize = 9, fontName = "Arial", fontColour = "red")
  expect_equal(
    getBaseFont(wb), 
    list(
      size = list(val = "9"),
      colour = list(rgb = "FFFF0000"),
      name = list(val = "Arial")
    )
  )
})
