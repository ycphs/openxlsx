
context("Testing 'topN' and 'bottomN' conditions in conditionalFormatting")
TNBN_test_data <- data.frame(col1 = 1:10,
                             col2 = 1:10,
                             col3 = seq(10, 100, 10),
                             col4 = seq(10, 100, 10),
                             col5 = 1:10,
                             col6 = 1:10)

bg_blue <- createStyle(bgFill = "skyblue")

wb <- createWorkbook()
sht <- "TopN_BottomN_TEST"
addWorksheet(wb, sht)
writeData(wb, sht, TNBN_test_data)
conditionalFormatting(wb, sht, cols = 1, rows = 2:11, type = "topN",    rank = 3,  style = bg_blue, percent = FALSE)
conditionalFormatting(wb, sht, cols = 2, rows = 2:11, type = "bottomN", rank = 3,  style = bg_blue, percent = FALSE)
conditionalFormatting(wb, sht, cols = 3, rows = 2:11, type = "topN",    rank = 50, style = bg_blue, percent = TRUE)
conditionalFormatting(wb, sht, cols = 4, rows = 2:11, type = "bottomN", rank = 50, style = bg_blue, percent = TRUE)
conditionalFormatting(wb, sht, cols = 5, rows = 2:11, type = "topN",    rank = 3,  style = bg_blue)
conditionalFormatting(wb, sht, cols = 6, rows = 2:11, type = "bottomN", rank = 3,  style = bg_blue)

test_that("Number of conditionalFormatting rules added equal to 6", {
  expect_equal(object = length(wb$worksheets[[1]]$conditionalFormatting), expected = 6)
})

test_that("topN conditions do not have the 'bottom' argument", {
  expect_false(object = grepl(paste('bottom'), wb$worksheets[[1]]$conditionalFormatting[1]))
  expect_false(object = grepl(paste('bottom'), wb$worksheets[[1]]$conditionalFormatting[3]))
})

test_that("bottomN conditions have the 'bottom' argument set to '1'", {
  expect_true(object = grepl(paste('bottom="1"'), wb$worksheets[[1]]$conditionalFormatting[2]))
  expect_true(object = grepl(paste('bottom="1"'), wb$worksheets[[1]]$conditionalFormatting[4]))
})

test_that("topN/bottomN rank conditions have the 'percent=FALSE' argument set to '0'", {
  expect_true(object = grepl(paste('percent="0"'), wb$worksheets[[1]]$conditionalFormatting[1]))
  expect_true(object = grepl(paste('percent="0"'), wb$worksheets[[1]]$conditionalFormatting[2]))
})

test_that("topN/bottomN rank conditions do not have the 'percent' argument is set to 'NULL'", {
  expect_true(object = grepl(paste('percent="NULL"'), wb$worksheets[[1]]$conditionalFormatting[5]))
  expect_true(object = grepl(paste('percent="NULL"'), wb$worksheets[[1]]$conditionalFormatting[6]))
})

test_that("topN/bottomN percent conditions have the 'percent' argument set to '1'", {
  expect_true(object = grepl(paste('percent="1"'), wb$worksheets[[1]]$conditionalFormatting[3]))
  expect_true(object = grepl(paste('percent="1"'), wb$worksheets[[1]]$conditionalFormatting[4]))
})

test_that("topN/bottomN conditions correspond to 'top10' type", {
  expect_true(object = grepl(paste('type="top10"'), wb$worksheets[[1]]$conditionalFormatting[1]))
  expect_true(object = grepl(paste('type="top10"'), wb$worksheets[[1]]$conditionalFormatting[2]))
  expect_true(object = grepl(paste('type="top10"'), wb$worksheets[[1]]$conditionalFormatting[3]))
  expect_true(object = grepl(paste('type="top10"'), wb$worksheets[[1]]$conditionalFormatting[4]))
  expect_true(object = grepl(paste('type="top10"'), wb$worksheets[[1]]$conditionalFormatting[5]))
  expect_true(object = grepl(paste('type="top10"'), wb$worksheets[[1]]$conditionalFormatting[6]))
})


context("Testing 'blanks' and 'notBlanks' conditions in conditionalFormatting")
BNB_test_data <- data.frame(col1 = sample(c("X", NA_character_), 10, replace = TRUE),
                             col2 = sample(c("Y", NA_character_), 10, replace = TRUE))

bg_blue <- createStyle(bgFill = "skyblue")
bg_red <- createStyle(bgFill = "red")

wb <- createWorkbook()
sht <- "Blanks_NonBlanks_TEST"
addWorksheet(wb, sht)
writeData(wb, sht, BNB_test_data)
conditionalFormatting(wb, sht, cols = 1, rows = 2:11, type = "blanks",    style = bg_red)
conditionalFormatting(wb, sht, cols = 2, rows = 2:11, type = "notBlanks", style = bg_blue)

test_that("Number of conditionalFormatting rules added equal to 2", {
  expect_equal(object = length(wb$worksheets[[1]]$conditionalFormatting), expected = 2)
})

test_that("type='blanks' calls type='containsBlanks'", {
  expect_true(object = grepl(paste('containsBlanks'), wb$worksheets[[1]]$conditionalFormatting[1]))
})

test_that("type='notBlanks' calls type='notContainsBlanks'", {
  expect_true(object = grepl(paste('notContainsBlanks'), wb$worksheets[[1]]$conditionalFormatting[2]))
})
