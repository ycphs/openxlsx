
#' @include class_definitions.R
Workbook$methods(writeData = function(
  df,
  sheet,
  startRow,
  startCol,
  colNames,
  colClasses,
  hlinkNames,
  keepNA, 
  na.string,
  list_sep
) {
  sheet <- validateSheet(sheet)
  nCols <- ncol(df)
  nRows <- nrow(df)
  df_nms <- names(df)

  allColClasses <- unlist(colClasses)
  
  isPOSIXlt <- function(data) sapply(lapply(data, class), FUN = function(x) any(x == "POSIXlt"))
  to_convert <- isPOSIXlt(df)
  
  if (any(to_convert)) {
    message("Found POSIXlt. Converting to POSIXct")
    df[to_convert] <- lapply(df[to_convert], as.POSIXct)  
  }
  
  
  df <- as.list(df)

  ######################################################################
  ## standardise all column types

  ## pull out NaN values
  nans <- unlist(lapply(1:nCols, function(i) {
    tmp <- df[[i]]
    if (!"character" %in% class(tmp) & !"list" %in% class(tmp)) {
      v <- which(is.nan(tmp) | is.infinite(tmp))
      if (length(v) == 0) {
        return(v)
      }
      return(as.integer(nCols * (v - 1) + i)) ## row position
    }
  }))

  ## convert any Dates to integers and create date style object
  if (any(c("date", "posixct", "posixt") %in% allColClasses)) {
    dInds <- which(sapply(colClasses, function(x) "date" %in% x))

    origin <- 25569L
    if (grepl('date1904="1"|date1904="true"', stri_join(unlist(workbook), collapse = ""), ignore.case = TRUE)) {
      origin <- 24107L
    }
    
    for (i in dInds) {
      df[[i]] <- as.integer(df[[i]]) + origin
      if (origin == 25569L) {
        earlyDate <- which(df[[i]] < 60)
        df[[i]][earlyDate] <- df[[i]][earlyDate] - 1
      }
    }

    pInds <- which(sapply(colClasses, function(x) any(c("posixct", "posixt", "posixlt") %in% x)))
    if (length(pInds) > 0 & nRows > 0) {
      parseOffset <- function(tz) {
        suppressWarnings(
          ifelse(stri_sub(tz, 1, 1) == "+", 1L, -1L)
          * (as.integer(stri_sub(tz, 2, 3)) + as.integer(stri_sub(tz, 4, 5)) / 60) / 24
        )
      }

      t <- lapply(df[pInds], function(x) format(x, "%z"))
      offSet <- lapply(t, parseOffset)
      offSet <- lapply(offSet, function(x) ifelse(is.na(x), 0, x))

      for (i in seq_along(pInds)) {
        df[[pInds[i]]] <- as.numeric(as.POSIXct(df[[pInds[i]]])) / 86400 + origin + offSet[[i]]
      }
    }
  }

  ## convert any Dates to integers and create date style object
  if (any(c("currency", "accounting", "percentage", "3", "comma") %in% allColClasses)) {
    cInds <- which(sapply(colClasses, function(x) any(c("accounting", "currency", "percentage", "3", "comma") %in% tolower(x))))
    for (i in cInds) {
      df[[i]] <- as.numeric(gsub("[^0-9\\.-]", "", df[[i]], perl = TRUE))
    }
    class(df[[i]]) <- "numeric"
  }

  ## convert scientific
  if ("scientific" %in% allColClasses) {
    for (i in which(sapply(colClasses, function(x) "scientific" %in% x))) {
      class(df[[i]]) <- "numeric"
    }
  }

  ##
  if ("list" %in% allColClasses) {
    for (i in which(sapply(colClasses, function(x) "list" %in% x))) {
      # check for and replace NA
      df_i <- lapply(df[[i]], unlist)
      df_i <- lapply(df_i, function(x) {
        x[is.na(x)] <- na.string
        x
      })
      df[[i]] <- sapply(df_i, stri_join, collapse = list_sep)
    }
  }

  if ("hyperlink" %in% allColClasses) {
    for (i in which(sapply(colClasses, function(x) "hyperlink" %in% x))) {
      class(df[[i]]) <- "hyperlink"
    }
  }

  if (any(c("formula", "array_formula") %in% allColClasses)) {
    
    frm <- "formula"
    cls <- "openxlsx_formula"
    
    if ("array_formula" %in% allColClasses) {
      frm <- "array_formula"
      cls <- "openxlsx_array_formula"
    }
    
    for (i in which(sapply(colClasses, function(x) frm %in% x))) {
      df[[i]] <- replaceIllegalCharacters(as.character(df[[i]]))
      class(df[[i]]) <- cls
    }
  }

  colClasses <- sapply(df, function(x) tolower(class(x))[[1]]) ## by here all cols must have a single class only


  ## convert logicals (Excel stores logicals as 0 & 1)
  if ("logical" %in% allColClasses) {
    for (i in which(sapply(colClasses, function(x) "logical" %in% x))) {
      class(df[[i]]) <- "numeric"
    }
  }

  ## convert all numerics to character (this way preserves digits)
  if ("numeric" %in% colClasses) {
    for (i in which(sapply(colClasses, function(x) "numeric" %in% x))) {
      class(df[[i]]) <- "character"
    }
  }

  ## End standardise all column types
  ######################################################################


  ## cell types
  t <- build_cell_types_integer(classes = colClasses, n_rows = nRows)

  for (i in which(sapply(colClasses, function(x) !"character" %in% x & !"numeric" %in% x))) {
    df[[i]] <- as.character(df[[i]])
  }

  ## cell values
  v <- as.character(t(as.matrix(
    data.frame(df, stringsAsFactors = FALSE, check.names = FALSE, fix.empty.names = FALSE)
  )))
  
  
  if (keepNA) {
    if (is.null(na.string)) {
      t[is.na(v)] <- 4L
      v[is.na(v)] <- "#N/A"
    } else {
      t[is.na(v)] <- 1L
      v[is.na(v)] <- as.character(na.string)
    }
  } else {
    t[is.na(v)] <- as.integer(NA)
    v[is.na(v)] <- as.character(NA)
  }

  ## If any NaN values
  if (length(nans) > 0) {
    t[nans] <- 4L
    v[nans] <- "#NUM!"
  }


  # prepend column headers
  if (colNames) {
    t <- c(rep.int(1L, nCols), t)
    v <- c(df_nms, v)
    nRows <- nRows + 1L
  }


  ## Formulas
  f_in <- rep.int(as.character(NA), length(t))
  any_functions <- FALSE
  ref_cell <- paste0(int_2_cell_ref(startCol), startRow)

  if (any(c("openxlsx_formula", "openxlsx_array_formula") %in% colClasses)) {
    
    ## alter the elements of t where we have a formula to be "str"
    if ("openxlsx_formula" %in% colClasses) {
      formula_cols <- which(sapply(colClasses, function(x) "openxlsx_formula" %in% x, USE.NAMES = FALSE), useNames = FALSE)
      formula_strs <- stri_join("<f>", unlist(df[formula_cols], use.names = FALSE), "</f>")
    } else { # openxlsx_array_formula
      formula_cols <- which(sapply(colClasses, function(x) "openxlsx_array_formula" %in% x, USE.NAMES = FALSE), useNames = FALSE)
      formula_strs <- stri_join("<f t=\"array\" ref=\"", ref_cell, ":", ref_cell, "\">", unlist(df[formula_cols], use.names = FALSE), "</f>")
    }
    formula_inds <- unlist(lapply(formula_cols, function(i) i + (1:(nRows - colNames) - 1) * nCols + (colNames * nCols)), use.names = FALSE)
    f_in[formula_inds] <- formula_strs
    any_functions <- TRUE

    rm(formula_cols)
    rm(formula_strs)
    rm(formula_inds)
  }

  suppressWarnings(try(rm(df), silent = TRUE))

  ## Append hyperlinks, convert h to s in cell type
  hyperlink_cols <- which(sapply(colClasses, function(x) "hyperlink" %in% x, USE.NAMES = FALSE), useNames = FALSE)
  if (length(hyperlink_cols) > 0) {
    hyperlink_inds <- sort(unlist(lapply(hyperlink_cols, function(i) i + (1:(nRows - colNames) - 1) * nCols + (colNames * nCols)), use.names = FALSE))
    na_hyperlink <- intersect(hyperlink_inds, which(is.na(t)))

    if (length(hyperlink_inds) > 0) {
      t[t %in% 9] <- 1L ## set cell type to "s"

      hyperlink_refs <- convert_to_excel_ref_expand(cols = hyperlink_cols + startCol - 1, LETTERS = LETTERS, rows = as.character((startRow + colNames):(startRow + nRows - 1L)))

      if (length(na_hyperlink) > 0) {
        to_remove <- which(hyperlink_inds %in% na_hyperlink)
        hyperlink_refs <- hyperlink_refs[-to_remove]
        hyperlink_inds <- hyperlink_inds[-to_remove]
      }

      exHlinks <- worksheets[[sheet]]$hyperlinks
      targets <- replaceIllegalCharacters(v[hyperlink_inds])

      if (!is.null(hlinkNames) & length(hlinkNames) == length(hyperlink_inds)) {
        v[hyperlink_inds] <- hlinkNames
      } ## this is text to display instead of hyperlink

      ## create hyperlink objects
      newhl <- lapply(seq_along(hyperlink_inds), function(i) {
        Hyperlink$new(ref = hyperlink_refs[i], target = targets[i], location = NULL, display = NULL, is_external = TRUE)
      })

      worksheets[[sheet]]$hyperlinks <<- append(worksheets[[sheet]]$hyperlinks, newhl)
    }
  }


  ## convert all strings to references in sharedStrings and update values (v)
  strFlag <- which(t == 1L)
  newStrs <- v[strFlag]
  if (length(newStrs) > 0) {
    newStrs <- replaceIllegalCharacters(newStrs)
    vl <- stri_length(newStrs)
    
    for (i in which(vl > 32767)) {
      
      if (vl[i]>32768+30) {
        warning(
          paste0(
            stri_sub(newStrs[i], 32768, 32768 + 15),
            " ... " ,
            stri_sub(newStrs[i], vl[i] - 15, vl[i]),
            " is truncated. 
Number of characters exeed the limit of 32767."
          )
        )
      } else {
        warning(
          paste0(
            stri_sub(newStrs[i], 32768, -1),
            " is truncated. 
Number of characters exeed the limit of 32767."
          )
        )
        
      }
      
      # v[i] <- stri_sub(v[i], 1, 32767)
    }
    newStrs <- stri_join("<si><t xml:space=\"preserve\">", newStrs, "</t></si>")

    uNewStr <- unique(newStrs)

    .self$updateSharedStrings(uNewStr)
    v[strFlag] <- match(newStrs, sharedStrings) - 1L
  }

  # ## Create cell list of lists
  worksheets[[sheet]]$sheet_data$write(
    rows_in = startRow:(startRow + nRows - 1L),
    cols_in = startCol:(startCol + nCols - 1L),
    t_in = t,
    v_in = v,
    f_in = f_in,
    any_functions = any_functions
  )

  invisible(0)
})
