test_that("createComment() works", {
  # error checking
  expect_error(createComment("hi", width = 1), NA)
  expect_error(createComment("hi", width = 1L), NA)
  expect_error(createComment("hi", width = 1:2), "width")
  
  expect_error(createComment("hi", height = 1), NA)
  expect_error(createComment("hi", height = 1L), NA)
  expect_error(createComment("hi", height = 1:2), "height")
  
  expect_error(createComment("hi", visible = NULL))
  expect_error(createComment("hi", visible = c(TRUE, FALSE)), "visible")
  
  expect_error(createComment("hi", author = 1))
  expect_error(createComment("hi", author = c("a", "a")), "author")
  
  expect_s4_class(createComment("hello"), "Comment")
})
