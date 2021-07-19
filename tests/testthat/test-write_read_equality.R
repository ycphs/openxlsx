

context("Writing and reading returns similar objects")

test_that("Writing then reading returns identical data.frame 1", {
  curr_wd <- getwd()

  ## data
  genDf <- function() {
    set.seed(1)
    data.frame(
      "Date" = Sys.Date() - 0:4,
      "Logical" = c(TRUE, FALSE, TRUE, TRUE, FALSE),
      "Currency" = -2:2,
      "Accounting" = -2:2,
      "hLink" = "https://CRAN.R-project.org/",
      "Percentage" = seq(-1, 1, length.out = 5),
      "TinyNumber" = runif(5) / 1E9, stringsAsFactors = FALSE
    )
  }

  df <- genDf()
  df

  class(df$Currency) <- "currency"
  class(df$Accounting) <- "accounting"
  class(df$hLink) <- "hyperlink"
  class(df$Percentage) <- "percentage"
  class(df$TinyNumber) <- "scientific"

  op <- options()
  options("openxlsx.dateFormat" = "yyyy-mm-dd")

  fileName <- file.path(tempdir(), "allClasses.xlsx")
  write.xlsx(df, file = fileName, overwrite = TRUE)

  x <- read.xlsx(xlsxFile = fileName, detectDates = TRUE)
  expect_equal(object = x, expected = genDf(), check.attributes = FALSE)

  unlink(fileName, recursive = TRUE, force = TRUE)

  expect_equal(object = getwd(), curr_wd)
  options(op)
})

test_that("Writing then reading returns identical data.frame 2", {
  curr_wd <- getwd()

  ## data.frame of dates
  dates <- data.frame("d1" = Sys.Date() - 0:500)
  dates[nrow(dates)+1,] = as.Date("1900-01-02")
  for (i in 1:3) dates <- cbind(dates, dates)
  names(dates) <- paste0("d", 1:8)

  ## Date Formatting
  wb <- createWorkbook()
  addWorksheet(wb, "Date Formatting", gridLines = FALSE)
  writeData(wb, 1, dates) ## write without styling

  ## set default date format
  options("openxlsx.dateFormat" = "yyyy/mm/dd")

  ## numFmt == "DATE" will use the date format specified by the above
  addStyle(wb, 1, style = createStyle(numFmt = "DATE"), rows = 2:11, cols = 1, gridExpand = TRUE)

  ## some custom date format examples
  sty <- createStyle(numFmt = "yyyy/mm/dd")
  addStyle(wb, 1, style = sty, rows = 2:11, cols = 2, gridExpand = TRUE)

  sty <- createStyle(numFmt = "yyyy/mmm/dd")
  addStyle(wb, 1, style = sty, rows = 2:11, cols = 3, gridExpand = TRUE)

  sty <- createStyle(numFmt = "yy / mmmm / dd")
  addStyle(wb, 1, style = sty, rows = 2:11, cols = 4, gridExpand = TRUE)

  sty <- createStyle(numFmt = "ddddd")
  addStyle(wb, 1, style = sty, rows = 2:11, cols = 5, gridExpand = TRUE)

  sty <- createStyle(numFmt = "yyyy-mmm-dd")
  addStyle(wb, 1, style = sty, rows = 2:11, cols = 6, gridExpand = TRUE)

  sty <- createStyle(numFmt = "mm/ dd yyyy")
  addStyle(wb, 1, style = sty, rows = 2:11, cols = 7, gridExpand = TRUE)

  sty <- createStyle(numFmt = "mm/dd/yy")
  addStyle(wb, 1, style = sty, rows = 2:11, cols = 8, gridExpand = TRUE)

  setColWidths(wb, 1, cols = 1:10, widths = 23)


  fileName <- file.path(tempdir(), "DateFormatting.xlsx")
  write.xlsx(dates, file = fileName, overwrite = TRUE)

  x <- read.xlsx(xlsxFile = fileName, detectDates = TRUE)
  expect_equal(object = x, expected = dates, check.attributes = FALSE)


  xNoDateDetection <- read.xlsx(xlsxFile = fileName, detectDates = FALSE)
  dateOrigin <- getDateOrigin(fileName)


  expect_equal(object = dateOrigin, expected = "1900-01-01", check.attributes = FALSE)

  for (i in seq_len(ncol(x))) {
    xNoDateDetection[[i]] <- convertToDate(xNoDateDetection[[i]], origin = dateOrigin)
  }

  expect_equal(object = xNoDateDetection, expected = dates, check.attributes = FALSE)

  expect_equal(object = getwd(), curr_wd)
  unlink(fileName, recursive = TRUE, force = TRUE)
})

test_that("Writing then reading rowNames, colNames combinations", {
  op <- options()
  options(stringsAsFactors = FALSE)
  
  
  fileName <- temp_xlsx()
  curr_wd <- getwd()
  mt <- utils::head(mtcars) # don't need the whole thing

  # write the row and column names for testing
  write.xlsx(mt, file = fileName, overwrite = TRUE, rowNames = TRUE, colNames = TRUE)
  
  # rowNames = colNames = TRUE
  # Row names = first column
  # Col names = first row
  x <- read.xlsx(fileName, sheet = 1, rowNames = TRUE)
  expect_equal(x, mt)


  # rowNames = TRUE, colNames = FALSE
  # Row names = first column
  # Col names = X1, X2, etc
  
  # need to create an expected output
  y <- as.data.frame(rbind(colnames(mt), as.matrix(mt)))
  colnames(y) <- c(make.names(seq_along(mt)))
  x <- read.xlsx(fileName, sheet = 1, rowNames = TRUE, colNames = FALSE)
  expect_equal(x, y)


  # rowNames = FALSE, colNames = TRUE
  # Row names = ""
  # Cl names = first row
  y2 <- cbind(row.names(mt), mt)
  colnames(y2)[1] <- ""
  row.names(y2) <- NULL
  x <- read.xlsx(fileName, sheet = 1, rowNames = FALSE, colNames = TRUE)
  expect_equal(x, y2)
  
  # rowNames = FALSE, colNames = FALSE
  # Row names = ""
  # Col names = X1, X2, etc
  y3 <- cbind(row.names(y), y)
  colnames(y3) <- make.names(seq_along(y3))
  row.names(y3) <- NULL
  x <- read.xlsx(fileName, sheet = 1, rowNames = FALSE, colNames = FALSE)
  expect_equal(x, y3)

  
  # Check wd
  expect_equal(getwd(), curr_wd)
  
  unlink(fileName, recursive = TRUE, force = TRUE)
  options(op)
})


test_that("Writing then reading returns identical data.frame 3", {
  op <- options()
  options(openxlsx.dateFormat = "yyyy-mm-dd")

  ## data
  df <- data.frame(
    Date       = as.Date("2021-05-21") - 0:4,
    Logical    = c(TRUE, FALSE, TRUE, TRUE, FALSE),
    Currency   = -2:2,
    Accounting = -2:2,
    hLink      = "https://CRAN.R-project.org/",
    Percentage = seq.int(-1, 1, length.out = 5),
    TinyNumber = runif(5) / 1E9,
    stringsAsFactors = FALSE
  )

  class(df$Currency) <- "currency"
  class(df$Accounting) <- "accounting"
  class(df$hLink) <- "hyperlink"
  class(df$Percentage) <- "percentage"
  class(df$TinyNumber) <- "scientific"

  fileName <- tempfile("allClasses", fileext = ".xlsx")
  write.xlsx(df, file = fileName, overwrite = TRUE)

  ## rows, cols combinations
  rows <- 1:4
  cols <- c(1, 3, 5)
  x <- read.xlsx(fileName, detectDates = TRUE, rows = rows, cols = cols)
  exp <- df[sort((rows - 1)[(rows - 1) <= nrow(df)]), sort(cols[cols <= ncol(df)])]
  expect_equal(x, exp)

  rows <- 1:4
  cols <- 1:9
  x <- read.xlsx(xlsxFile = fileName, detectDates = TRUE, rows = rows, cols = cols)
  exp <- df[sort((rows - 1)[(rows - 1) <= nrow(df)]), sort(cols[cols <= ncol(df)])]
  expect_equal(x, exp)

  rows <- 1:200
  cols <- c(5, 99, 2)
  x <- read.xlsx(xlsxFile = fileName, detectDates = TRUE, rows = rows, cols = cols)
  exp <- df[sort((rows - 1)[(rows - 1) <= nrow(df)]), sort(cols[cols <= ncol(df)])]
  expect_equal(x, exp)


  rows <- 1000:900
  cols <- c(5, 99, 2)
  suppressWarnings(x <- read.xlsx(xlsxFile = fileName, detectDates = TRUE, rows = rows, cols = cols))
  expect_identical(x, NULL)

  unlink(fileName, recursive = TRUE, force = TRUE)
  options(op)
})


test_that("Writing then reading returns identical data.frame 4", {

  ## data
  df <- head(iris[, 1:4])
  df[1, 2] <- NA
  df[3, 1] <- NA
  df[6, 4] <- NA


  tf <- temp_xlsx()
  write.xlsx(x = df, file = tf, keepNA = TRUE)
  x <- read.xlsx(tf)

  expect_equal(object = x, expected = df, check.attributes = TRUE)
  unlink(tf, recursive = TRUE, force = TRUE)


  tf <- temp_xlsx()
  write.xlsx(x = df, file = tf, keepNA = FALSE)
  x <- read.xlsx(tf)

  expect_equal(object = x, expected = df, check.attributes = TRUE)
  unlink(tf, recursive = TRUE, force = TRUE)
})

test_that("Writing then reading returns identical data.frame 5", {

  ## data
  df <- head(iris[, 1:4])
  df[1, 2] <- NA
  df[3, 1] <- NA
  df[6, 4] <- NA

  na.string <- "*"
  df_expected <- df
  df_expected[1, 2] <- na.string
  df_expected[3, 1] <- na.string
  df_expected[6, 4] <- na.string


  tf <- temp_xlsx()
  write.xlsx(x = df, file = tf, keepNA = TRUE, na.string = na.string)
  x <- read.xlsx(tf)

  expect_equal(object = x, expected = df_expected, check.attributes = TRUE)
  unlink(tf, recursive = TRUE, force = TRUE)
})



test_that("Special characters in sheet names", {
  tf <- temp_xlsx()

  ## data
  sheet_name <- "A & B < D > D"

  wb <- createWorkbook()
  addWorksheet(wb, sheetName = sheet_name)
  addWorksheet(wb, sheetName = "test")
  writeData(wb, sheet = 1, x = 1:10)
  saveWorkbook(wb = wb, file = tf, overwrite = TRUE)

  expect_equal(getSheetNames(tf)[1], sheet_name)
  expect_equal(getSheetNames(tf)[2], "test")

  expect_equal(read.xlsx(tf, colNames = FALSE)[[1]], 1:10)

  unlink(tf, recursive = TRUE, force = TRUE)
})
