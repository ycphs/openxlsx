

Comment <- setRefClass("Comment",
  fields = c(
    "text",
    "author",
    "style",
    "visible",
    "width",
    "height"
  ),

  methods = list()
)


Comment$methods(initialize = function(text, author, style, visible = TRUE, width = 2, height = 4) {
  text <<- text
  author <<- author
  style <<- style
  visible <<- visible
  width <<- width
  height <<- height
})


Comment$methods(show = function() {
  showText <- sprintf("Author: %s\n", author)
  showText <- c(showText, sprintf("Text:\n %s\n\n", paste(text, collapse = "")))
  styleShow <- "Style:\n"

  if ("list" %in% class(style)) {
    for (i in seq_along(style)) {
      styleShow <- append(styleShow, sprintf("Font name: %s\n", style[[i]]$fontName[[1]])) ## Font name
      styleShow <- append(styleShow, sprintf("Font size: %s\n", style[[i]]$fontSize[[1]])) ## Font size
      styleShow <- append(styleShow, sprintf("Font colour: %s\n", gsub("^FF", "#", style[[i]]$fontColour[[1]]))) ## Font colour

      ## Font decoration
      if (length(style[[i]]$fontDecoration) > 0) {
        styleShow <- append(styleShow, sprintf("Font decoration: %s\n", paste(style[[i]]$fontDecoration, collapse = ", ")))
      }

      styleShow <- append(styleShow, "\n\n")
    }
  } else {
    styleShow <- append(styleShow, sprintf("Font name: %s \n", style$fontName[[1]])) ## Font name
    styleShow <- append(styleShow, sprintf("Font size: %s \n", style$fontSize[[1]])) ## Font size
    styleShow <- append(styleShow, sprintf("Font colour: %s \n", gsub("^FF", "#", style$fontColour[[1]]))) ## Font colour

    ## Font decoration
    if (length(style$fontDecoration) > 0) {
      styleShow <- append(styleShow, sprintf("Font decoration: %s \n", paste(style$fontDecoration, collapse = ", ")))
    }

    styleShow <- append(styleShow, "\n\n")
  }

  showText <- paste0(paste(showText, collapse = ""), paste(styleShow, collapse = ""), collapse = "")
  cat(showText)
})



#' @name createComment
#' @title create a Comment object
#' @description Create a cell Comment object to pass to writeComment()
#' @param comment Comment text. Character vector.
#' @param author Author of comment. Character vector of length 1
#' @param style A Style object or list of style objects the same length as comment vector. See [createStyle()].
#' @param visible TRUE or FALSE. Is comment visible.
#' @param width,height Width and height of textbook (in number of cells);
#'   doubles are rounded with \code{base::round()}
#' @export
#' @seealso [writeComment()]
#' @examples
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet 1")
#'
#' c1 <- createComment(comment = "this is comment")
#' writeComment(wb, 1, col = "B", row = 10, comment = c1)
#'
#' s1 <- createStyle(fontSize = 12, fontColour = "red", textDecoration = c("BOLD"))
#' s2 <- createStyle(fontSize = 9, fontColour = "black")
#'
#' c2 <- createComment(comment = c("This Part Bold red\n\n", "This part black"), style = c(s1, s2))
#' c2
#'
#' writeComment(wb, 1, col = 6, row = 3, comment = c2)
#' \dontrun{
#' saveWorkbook(wb, file = "createCommentExample.xlsx", overwrite = TRUE)
#' }
createComment <- function(comment,
                          author = Sys.getenv("USERNAME"),
                          style = NULL,
                          visible = TRUE,
                          width = 2,
                          height = 4) {
  
  if (!is.character(comment)) {
    stop("comment argument must be a character vector")
  }

  assert_character1(author)
  assert_numeric1(width)
  assert_numeric1(height)
  assert_true_false1(visible)

  width <- round(width)
  height <- round(height)

  if (is.null(style)) {
    style <- createStyle(fontName = "Tahoma", fontSize = 9, fontColour = "black")
  }

  author <- replaceIllegalCharacters(author)
  comment <- replaceIllegalCharacters(comment)

  invisible(Comment$new(text = comment, author = author, style = style, visible = visible, width = width[1], height = height[1]))
}


#' @name writeComment
#' @title write a cell comment
#' @description Write a Comment object to a worksheet
#' @param wb A workbook object
#' @param sheet A vector of names or indices of worksheets
#' @param col Column a column number of letter
#' @param row A row number.
#' @param comment A Comment object. See [createComment()].
#' @param xy An alternative to specifying `col` and
#' `row` individually.  A vector of the form
#' `c(col, row)`.
#' @export
#' @seealso [createComment()]
#' @examples
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet 1")
#'
#' c1 <- createComment(comment = "this is comment")
#' writeComment(wb, 1, col = "B", row = 10, comment = c1)
#'
#' s1 <- createStyle(fontSize = 12, fontColour = "red", textDecoration = c("BOLD"))
#' s2 <- createStyle(fontSize = 9, fontColour = "black")
#'
#' c2 <- createComment(comment = c("This Part Bold red\n\n", "This part black"), style = c(s1, s2))
#' c2
#'
#' writeComment(wb, 1, col = 6, row = 3, comment = c2)
#' \dontrun{
#' saveWorkbook(wb, file = "writeCommentExample.xlsx", overwrite = TRUE)
#' }
writeComment <- function(wb, sheet, col, row, comment, xy = NULL) {
  if (!"Workbook" %in% class(wb)) {
    stop("First argument must be a Workbook.")
  }

  if (!"Comment" %in% class(comment)) {
    stop("comment argument must be a Comment object")
  }


  if (length(comment$style) == 1) {
    rPr <- wb$createFontNode(comment$style)
  } else {
    rPr <- sapply(comment$style, function(x) wb$createFontNode(x))
  }

  rPr <- gsub("font>", "rPr>", rPr)
  sheet <- wb$validateSheet(sheet)

  ## All input conversions/validations
  if (!is.null(xy)) {
    if (length(xy) != 2) {
      stop("xy parameter must have length 2")
    }
    col <- xy[[1]]
    row <- xy[[2]]
  }

  if (!is.numeric(col)) {
    col <- convertFromExcelRef(col)
  }

  ref <- paste0(convert_to_excel_ref(cols = col, LETTERS = LETTERS), row)

  comment_list <- list(
    "ref" = ref,
    "author" = comment$author,
    "comment" = comment$text,
    "style" = rPr,
    "clientData" = genClientData(col, row, visible = comment$visible, height = comment$height, width = comment$width)
  )

  wb$comments[[sheet]] <- append(wb$comments[[sheet]], list(comment_list))

  invisible(wb)
}


#' @name removeComment
#' @title Remove a comment from a cell
#' @description Remove a cell comment from a worksheet
#' @param wb A workbook object
#' @param sheet A vector of names or indices of worksheets
#' @param cols Columns to delete comments from
#' @param rows Rows to delete comments from
#' @param gridExpand If `TRUE`, all data in rectangle min(rows):max(rows) X min(cols):max(cols)
#' will be removed.
#' @export
#' @seealso [createComment()]
#' @seealso [writeComment()]
removeComment <- function(wb, sheet, cols, rows, gridExpand = TRUE) {
  sheet <- wb$validateSheet(sheet)

  assert_class(wb, "Workbook")
  cols <- convertFromExcelRef(cols)
  rows <- as.integer(rows)

  ## rows and cols need to be the same length
  if (gridExpand) {
    combs <- expand.grid(rows, cols)
    rows <- combs[, 1]
    cols <- combs[, 2]
  }

  if (length(rows) != length(cols)) {
    stop("Length of rows and cols must be equal.")
  }

  comb <- paste0(convert_to_excel_ref(cols = cols, LETTERS = LETTERS), rows)
  toKeep <- !sapply(wb$comments[[sheet]], "[[", "ref") %in% comb

  wb$comments[[sheet]] <- wb$comments[[sheet]][toKeep]
}
