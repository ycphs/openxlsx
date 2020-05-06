
context("Workbook properties")

test_that("Workbook properties", {

  ## check creator
  wb <- createWorkbook(creator = "Alex", title = "title here", subject = "this & that", category = "some category")

  expect_true(grepl("<dc:creator>Alex</dc:creator>", wb$core))
  expect_true(grepl("<dc:title>title here</dc:title>", wb$core))
  expect_true(grepl("<dc:subject>this &amp; that</dc:subject>", wb$core))
  expect_true(grepl("<cp:category>some category</cp:category>", wb$core))

  fl <- tempfile(fileext = ".xlsx")
  wb <- write.xlsx(
    x = iris, file = fl,
    creator = "Alex 2",
    title = "title here 2",
    subject = "this & that 2",
    category = "some category 2"
  )

  expect_true(grepl("<dc:creator>Alex 2</dc:creator>", wb$core))
  expect_true(grepl("<dc:title>title here 2</dc:title>", wb$core))
  expect_true(grepl("<dc:subject>this &amp; that 2</dc:subject>", wb$core))
  expect_true(grepl("<cp:category>some category 2</cp:category>", wb$core))

  ## maintain on load
  wb_loaded <- loadWorkbook(fl)
  expect_equal(object = wb_loaded$core, expected = paste0(wb$core, collapse = ""))


  wb <- createWorkbook(creator = "Philipp", title = "title here", subject = "this & that", category = "some category")
  addCreator(wb, "test")
  expect_true(grepl("<dc:creator>Philipp;test</dc:creator>", wb$core))

  expect_equal(getCreators(wb), c("Philipp", "test"))
  setLastModifiedBy(wb, "Philipp 2")
  expect_true(grepl("<cp:lastModifiedBy>Philipp 2</cp:lastModifiedBy>", wb$core))
})
