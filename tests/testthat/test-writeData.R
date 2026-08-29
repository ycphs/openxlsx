test_that("writeData() forces evaluation of x (#264)", {
  wbfile <- temp_xlsx()
  op <- options(stringsAsFactors = FALSE)
  
  x <- format(123.4)
  df <- data.frame(d = format(123.4))
  df2 <- data.frame(e = x)
  
  wb <- createWorkbook()
  addWorksheet(wb, "sheet")
  writeData(wb, "sheet", startCol = 1, data.frame(a = format(123.4)))
  writeData(wb, "sheet", startCol = 2, data.frame(b = as.character(123.4)))
  writeData(wb, "sheet", startCol = 3, data.frame(c = "123.4"))
  writeData(wb, "sheet", startCol = 4, df)
  writeData(wb, "sheet", startCol = 5, df2)
  
  saveWorkbook(wb, wbfile)
  out <- read.xlsx(wbfile)
  
  # Possibly overkill
  
  with(out, {
    expect_identical(a, b)
    expect_identical(a, c)
    expect_identical(a, d)
    expect_identical(a, e)
    expect_identical(b, c)
    expect_identical(b, d)
    expect_identical(b, e)
    expect_identical(c, d)
    expect_identical(c, e)
    expect_identical(d, e)
  })
  
  options(op)
  file.remove(wbfile)
})


test_that("as.character.formula() works [312]", {
  form <- y ~ x1 * x2 + x3
  expect_identical(
    as.character.default(form),
    openxlsx:::as.character.formula(form)
  )
  
  skip_if_not_installed("formula.tools")
  # "tests specifically for as.character.formula conflict"
  
  foo <- function() {
    wb <- openxlsx::buildWorkbook(
      data.frame(
        x = structure("A2 + B2", class = c("character", "formula")),
        stringsAsFactors = FALSE
      )
    )
    as.list(wb$worksheets[[1]]$sheet_data)
  }
  
  before <- foo()
  # don't required the "require" function for deps  check
  match.fun("require")("formula.tools", character.only = TRUE)
  middle <- foo()
  detach("package:formula.tools", character.only = TRUE, force = TRUE)
  end <- foo()
  
  expect_identical(before, middle, ignore.environment = TRUE)
  expect_identical(before, end,    ignore.environment = TRUE)
})


test_that("writeData(overwrite = TRUE) works (#536)", {
  wb <- createWorkbook()
  addWorksheet(wb, "sheet")

  old <- data.frame(a = 1:6, b = letters[1:6], stringsAsFactors = FALSE)
  new <- old[1:4, , drop = FALSE]

  writeData(wb, "sheet", old)
  writeData(wb, "sheet", new)
  out_no_overwrite <- readWorkbook(wb, "sheet")
  expect_equal(nrow(out_no_overwrite), 6)

  writeData(wb, "sheet", new, overwrite = TRUE)
  out_overwrite <- readWorkbook(wb, "sheet")
  expect_equal(out_overwrite, new)

  wb2 <- createWorkbook()
  addWorksheet(wb2, "sheet")
  tab_df <- data.frame(x = 1:2)
  writeDataTable(wb2, "sheet", tab_df)
  before <- readWorkbook(wb2, "sheet")

  expect_error(
    writeData(wb2, "sheet", data.frame(y = 1), overwrite = TRUE),
    "Cannot overwrite table headers"
  )

  after <- readWorkbook(wb2, "sheet")
  expect_equal(after, before)
})
