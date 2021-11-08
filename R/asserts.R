# Assertions for parameter validates
# These should be used at the beginning of functions to stop execution early

assert_class <- function(x, class, or_null = FALSE) {
  sx <- as.character(substitute(x))
  ok <- inherits(x, class)
  
  if (or_null) {
    ok <- ok | is.null(x)
    class <- c(class, "null")
  }
  
  if (!ok) {
    msg <- sprintf("%s must be of class %s", sx, paste(class, collapse = " or "))
    stop(msg, call. = FALSE)
  }
}

assert_length <- function(x, n) {
  stopifnot(is.integer(n))
  if (length(x) != n) {
    msg <- sprintf("%s must be of length %iL", substitute(x), n)
    stop(msg, call. = FALSE)
  }
}

assert_true_false1 <- function(x) {
  if (!is_true_false(x)) {
    stop(substitute(x), " must be TRUE or FALSE", call. = FALSE)
  }
}

assert_true_false <- function(x) {
  ok <- is.logical(x) & !is.na(x)
  if (!ok) {
    stop(substitute(x), " must be a logical vector with NAs", call. = FALSE)
  }
}

assert_character1 <- function(x, scalar = FALSE) {
  ok <- is.character(x) && length(x) == 1L
  
  if (scalar) {
    ok <- ok & nchar(x) == 1L
  }
  
  if (!ok) {
    stop(substitute(x), " must be a character vector of length 1L", call. = FALSE)
  }
}

assert_unique <- function(x, case_sensitive = TRUE) {
  msg <- paste0(substitute(x), " must be a unique vector")
  
  if (!case_sensitive) {
    x <- tolower(x)
    msg <- paste0(msg, " (case sensitive)")
  }
  
  if (anyDuplicated(x) != 0L) {
    stop(msg, call. = FALSE)
  }
}

assert_numeric1 <- function(x, scalar = FALSE) {
  msg <- paste0(substitute(x), " must be a ")
  ok <- is.numeric(x) & length(x) == 1L
  
  if (scalar) {
    ok <- ok && nchar(x) == 1L
    msg <- paste0(msg, "single number")
  } else {
    msg <- paste0(msg, "numeric vector of length 1L")
  }
  
  if (!ok) {
    stop(msg, call. = FALSE)
  }
}

# validates ---------------------------------------------------------------

validate_StyleName <- function(x) {
  m <- valid_StyleNames[match(tolower(x), valid_StyleNames_low)]
  if (anyNA(m)) {
    stop(
      "Invalid table style: ", 
      paste0(sprintf("'%s'", x[is.na(m)]), collapse = ", "),
      call. = FALSE
    )
  }
  m
}

valid_StyleNames <- c("none", paste0("TableStyleLight", 1:21), paste0("TableStyleMedium", 1:28), paste0("TableStyleDark", 1:11))
valid_StyleNames_low <- tolower(valid_StyleNames)
