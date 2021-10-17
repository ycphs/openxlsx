test_that("createComment() works", {
  # error checking
  expect_error(createComment(width = 1), NA)
  expect_error(createComment(width = 1L), NA)
  expect_error(createComment(width = 1:2))
  
  expect_error(createComment(height = 1), NA)
  expect_error(createComment(height = 1L), NA)
  expect_error(createComment(height = 1:2))
  
  expect_error(createComment(visible = NULL))
  expect_error(createComment(visible = c(TRUE, FALSE)))
  
  expect_error(createComment(author = 1))
  expect_error(createComment(author = c("a", "a")))
  
  expect_s4_class(createComment("hello"), "Comment")
})
