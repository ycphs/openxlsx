


#' @name loadWorkbook
#' @title Load an existing .xlsx file
#' @author Alexander Walker, Philipp Schauberger
#' @param file A path to an existing .xlsx or .xlsm file
#' @param xlsxFile alias for file
#' @param isUnzipped Set to TRUE if the xlsx file is already unzipped
#' @description  loadWorkbook returns a workbook object conserving styles and
#' formatting of the original .xlsx file.
#' @return Workbook object.
#' @export
#' @seealso [removeWorksheet()]
#' @examples
#' ## load existing workbook from package folder
#' wb <- loadWorkbook(file = system.file("extdata", "loadExample.xlsx", package = "openxlsx"))
#' names(wb) # list worksheets
#' wb ## view object
#' ## Add a worksheet
#' addWorksheet(wb, "A new worksheet")
#'
#' ## Save workbook
#' \dontrun{
#' saveWorkbook(wb, "loadExample.xlsx", overwrite = TRUE)
#' }
#'
loadWorkbook <- function(file, xlsxFile = NULL, isUnzipped = FALSE) {

  ## If this is a unzipped workbook, skip the temp dir stuff
  if (isUnzipped) {
    xmlDir <- file
    xmlFiles <- list.files(path = xmlDir, full.names = TRUE, recursive = TRUE, all.files = TRUE)
  } else {
    if (!is.null(xlsxFile)) {
      file <- xlsxFile
    }

    file <- getFile(file)
    if (!file.exists(file)) {
      stop("File does not exist.")
    }

    ## create temp dir
    xmlDir <- tempfile()

    ## Unzip files to temp directory
    xmlFiles <- unzip(file, exdir = xmlDir)
  }
  wb <- createWorkbook()

  ## Not used
  # .relsXML           <- xmlFiles[grepl("_rels/.rels$", xmlFiles, perl = TRUE)]
  # appXML             <- xmlFiles[grepl("app.xml$", xmlFiles, perl = TRUE)]

  drawingsXML   <- grep("drawings/drawing[0-9]+.xml$", xmlFiles, perl = TRUE, value = TRUE)
  worksheetsXML <- grep("/worksheets/sheet[0-9]+", xmlFiles, perl = TRUE, value = TRUE)

  coreXML           <- grep("core.xml$", xmlFiles, perl = TRUE, value = TRUE)
  workbookXML       <- grep("workbook.xml$", xmlFiles, perl = TRUE, value = TRUE)
  stylesXML         <- grep("styles.xml$", xmlFiles, perl = TRUE, value = TRUE)
  sharedStringsXML  <- grep("sharedStrings.xml$", xmlFiles, perl = TRUE, value = TRUE)
  themeXML          <- grep("theme[0-9]+.xml$", xmlFiles, perl = TRUE, value = TRUE)
  drawingRelsXML    <- grep("drawing[0-9]+.xml.rels$", xmlFiles, perl = TRUE, value = TRUE)
  sheetRelsXML      <- grep("sheet[0-9]+.xml.rels$", xmlFiles, perl = TRUE, value = TRUE)
  media             <- grep("image[0-9]+.[a-z]+$", xmlFiles, perl = TRUE, value = TRUE)
  vmlDrawingXML     <- grep("drawings/vmlDrawing[0-9]+\\.vml$", xmlFiles, perl = TRUE, value = TRUE)
  vmlDrawingRelsXML <- grep("vmlDrawing[0-9]+.vml.rels$", xmlFiles, perl = TRUE, value = TRUE)
  commentsXML       <- grep("xl/comments[0-9]+\\.xml", xmlFiles, perl = TRUE, value = TRUE)
  threadCommentsXML <- grep("xl/threadedComments/threadedComment[0-9]+\\.xml", xmlFiles, perl = TRUE, value = TRUE)
  personXML         <- grep("xl/persons/person.xml$", xmlFiles, perl = TRUE, value = TRUE)
  embeddings        <- grep("xl/embeddings", xmlFiles, perl = TRUE, value = TRUE)
  charts            <- grep("xl/charts/.*xml$", xmlFiles, perl = TRUE, value = TRUE)
  chartsRels        <- grep("xl/charts/_rels", xmlFiles, perl = TRUE, value = TRUE)
  chartSheetsXML    <- grep("xl/chartsheets/sheet[0-9]+\\.xml", xmlFiles, perl = TRUE, value = TRUE)
  tablesXML         <- grep("tables/table[0-9]+.xml$", xmlFiles, perl = TRUE, value = TRUE)
  tableRelsXML      <- grep("table[0-9]+.xml.rels$", xmlFiles, perl = TRUE, value = TRUE)
  queryTablesXML    <- grep("queryTable[0-9]+.xml$", xmlFiles, perl = TRUE, value = TRUE)
  connectionsXML    <- grep("connections.xml$", xmlFiles, perl = TRUE, value = TRUE)
  extLinksXML       <- grep("externalLink[0-9]+.xml$", xmlFiles, perl = TRUE, value = TRUE)
  extLinksRelsXML   <- grep("externalLink[0-9]+.xml.rels$", xmlFiles, perl = TRUE, value = TRUE)


  # pivot tables
  pivotTableXML     <- grep("pivotTable[0-9]+.xml$", xmlFiles, perl = TRUE, value = TRUE)
  pivotTableRelsXML <- grep("pivotTable[0-9]+.xml.rels$", xmlFiles, perl = TRUE, value = TRUE)
  pivotDefXML       <- grep("pivotCacheDefinition[0-9]+.xml$", xmlFiles, perl = TRUE, value = TRUE)
  pivotDefRelsXML   <- grep("pivotCacheDefinition[0-9]+.xml.rels$", xmlFiles, perl = TRUE, value = TRUE)
  pivotCacheRecords <- grep("pivotCacheRecords[0-9]+.xml$", xmlFiles, perl = TRUE, value = TRUE)

  ## slicers
  slicerXML       <- grep("slicer[0-9]+.xml$", xmlFiles, perl = TRUE, value = TRUE)
  slicerCachesXML <- grep("slicerCache[0-9]+.xml$", xmlFiles, perl = TRUE, value = TRUE)

  ## VBA Macro
  vbaProject <- grep("vbaProject\\.bin$", xmlFiles, perl = TRUE, value = TRUE)

  ## remove all EXCEPT media and charts
  if (!isUnzipped) {
    on.exit({
      paths <- grep(
        "charts|media|vmlDrawing|comment|embeddings|pivot|slicer|vbaProject|person", 
        xmlFiles, 
        ignore.case = TRUE,
        value = TRUE,
        invert = TRUE
      )
      unlink(paths, recursive = TRUE, force = TRUE)
    },
      add = TRUE
    )
}

  ## core
  if (length(coreXML) == 1) {
    coreXML <- paste(readUTF8(coreXML), collapse = "")
    wb$core <- removeHeadTag(x = coreXML)
  }

  nSheets <- length(worksheetsXML) + length(chartSheetsXML)

  ## get Rid of chartsheets, these do not have a worksheet/sheeti.xml
  worksheet_rId_mapping <- NULL
  workbookRelsXML <- grep("workbook.xml.rels$", xmlFiles, perl = TRUE, value = TRUE)
  if (length(workbookRelsXML) > 0) {
    workbookRelsXML <- paste(readUTF8(workbookRelsXML), collapse = "")
    workbookRelsXML <- getChildlessNode(xml = workbookRelsXML, tag = "Relationship")
    worksheet_rId_mapping <- grep("worksheets/sheet", workbookRelsXML, fixed = TRUE, value = TRUE)
  }

  ##
  chartSheetRIds <- NULL
  if (length(chartSheetsXML) > 0) {
    workbookRelsXML <- grep("chartsheets/sheet", workbookRelsXML, fixed = TRUE, value = TRUE)

    chartSheetRIds <- unlist(getId(workbookRelsXML))
    chartsheet_rId_mapping <- unlist(regmatches(workbookRelsXML, gregexpr("sheet[0-9]+\\.xml", workbookRelsXML, perl = TRUE, ignore.case = TRUE)))

    sheetNo <- as.integer(regmatches(chartSheetsXML, regexpr("(?<=sheet)[0-9]+(?=\\.xml)", chartSheetsXML, perl = TRUE)))
    chartSheetsXML <- chartSheetsXML[order(sheetNo)]

    chartSheetsRelsXML <- grep("xl/chartsheets/_rels", xmlFiles, perl = TRUE, value = TRUE)
    sheetNo2 <- as.integer(regmatches(chartSheetsRelsXML, regexpr("(?<=sheet)[0-9]+(?=\\.xml\\.rels)", chartSheetsRelsXML, perl = TRUE)))
    chartSheetsRelsXML <- chartSheetsRelsXML[order(sheetNo2)]

    chartSheetsRelsDir <- dirname(chartSheetsRelsXML[1])
  }


  ## xl\
  ## xl\workbook
  if (length(workbookXML) > 0) {
    workbook <- readUTF8(workbookXML)
    workbook <- removeHeadTag(workbook)

    sheets <- unlist(regmatches(workbook, gregexpr("(?<=<sheets>).*(?=</sheets>)", workbook, perl = TRUE)))
    sheets <- unlist(regmatches(sheets, gregexpr("<sheet[^>]*>", sheets, perl = TRUE)))
    
    ## Some veryHidden sheets do not have a sheet content and their rId is empty.
    ## Such sheets need to be filtered out because otherwise their sheet names
    ## occur in the list of all sheet names, leading to a wrong association
    ## of sheet names with sheet indeces.
    sheets <- grep('r:id="[[:blank:]]*"', sheets, invert = TRUE, value = TRUE)


    ## sheetId is meaningless
    ## sheet rId links to the workbook.xml.resl which links worksheets/sheet(i).xml file
    ## order they appear here gives order of worksheets in xlsx file

    sheetrId <- unlist(getRId(sheets))
    sheetId <- unlist(regmatches(sheets, gregexpr('(?<=sheetId=")[0-9]+', sheets, perl = TRUE)))
    sheetNames <- unlist(regmatches(sheets, gregexpr('(?<=name=")[^"]+', sheets, perl = TRUE)))
    sheetNames <- replaceXMLEntities(sheetNames)


    is_chart_sheet <- sheetrId %in% chartSheetRIds
    is_visible <- !grepl("hidden", sheets)
    if (length(is_visible) != length(sheetrId)) {
      is_visible <- rep(TRUE, length(sheetrId))
    }
    

# #active sheet -----------------------------------------------------------

      
   
    
    ## add worksheets to wb
    j <- 1
    for (i in seq_along(sheetrId)) {
      if (is_chart_sheet[i]) {
        # count <- 0  variable not used
        txt <- paste(readUTF8(chartSheetsXML[j]), collapse = "")

        zoom <- regmatches(txt, regexpr('(?<=zoomScale=")[0-9]+', txt, perl = TRUE))
        if (length(zoom) == 0) {
          zoom <- 100
        }

        tabColour <- getChildlessNode(xml = txt, tag = "tabColor")
        if (length(tabColour) == 0) {
          tabColour <- NULL
        }

        j <- j + 1L

        wb$addChartSheet(sheetName = sheetNames[i], tabColour = tabColour, zoom = as.numeric(zoom))
      } else {
        wb$addWorksheet(sheetNames[i], visible = is_visible[i])
      }
    }


    ## replace sheetId
    for (i in 1:nSheets) {
      wb$workbook$sheets[[i]] <- gsub(sprintf(' sheetId="%s"', i), sprintf(' sheetId="%s"', sheetId[i]), wb$workbook$sheets[[i]])
    }


    ## additional workbook attributes
    calcPr <- getChildlessNode(xml = workbook, tag = "calcPr")
    if (length(calcPr) > 0) {
      wb$workbook$calcPr <- calcPr
    }

    ## additional workbook attributes
    extLst <- getChildlessNode(xml = workbook, tag = "extLst")
    if (length(extLst) > 0) {
      wb$workbook$extLst <- extLst
    }

    workbookPr <- getChildlessNode(xml = workbook, tag = "workbookPr")
    if (length(workbookPr) > 0) {
      wb$workbook$workbookPr <- workbookPr
    }

    workbookProtection <- getChildlessNode(xml = workbook, tag = "workbookProtection")
    if (length(workbookProtection) > 0) {
      wb$workbook$workbookProtection <- workbookProtection
    }


    ## defined Names
    dNames <- getNodes(xml = workbook, tagIn = "<definedNames>")
    if (length(dNames) > 0) {
      dNames <- gsub("^<definedNames>|</definedNames>$", "", dNames)
      wb$workbook$definedNames <- paste0(getNodes(xml = dNames, tagIn = "<definedName"), ">")
    }
  }





  ## xl\sharedStrings
  if (length(sharedStringsXML) > 0) {
    sharedStrings <- readUTF8(sharedStringsXML)
    sharedStrings <- paste(sharedStrings, collapse = "\n")
    sharedStrings <- removeHeadTag(sharedStrings)

    uniqueCount <- as.integer(regmatches(sharedStrings, regexpr('(?<=uniqueCount=")[0-9]+', sharedStrings, perl = TRUE)))

    ## read in and get <si> nodes
    vals <- getNodes(xml = sharedStrings, tagIn = "<si>")

    if ("<si><t/></si>" %in% vals) {
      vals[vals == "<si><t/></si>"] <- "<si><t>NA</t></si>"
      Encoding(vals) <- "UTF-8"
      attr(vals, "uniqueCount") <- uniqueCount - 1L
    } else {
      Encoding(vals) <- "UTF-8"
      attr(vals, "uniqueCount") <- uniqueCount
    }

    wb$sharedStrings <- vals
  }

  ## xl\pivotTables & xl\pivotCache
  if (length(pivotTableXML) > 0) {

    # pivotTable cacheId links to workbook.xml which links to workbook.xml.rels via rId
    # we don't modify the cacheId, only the rId
    nPivotTables <- length(pivotDefXML)
    rIds <- 20000L + 1:nPivotTables

    ## pivot tables
    pivotTableXML <- pivotTableXML[order(nchar(pivotTableXML), pivotTableXML)]
    pivotTableRelsXML <- pivotTableRelsXML[order(nchar(pivotTableRelsXML), pivotTableRelsXML)]

    ## Cache
    pivotDefXML <- pivotDefXML[order(nchar(pivotDefXML), pivotDefXML)]
    pivotDefRelsXML <- pivotDefRelsXML[order(nchar(pivotDefRelsXML), pivotDefRelsXML)]
    pivotCacheRecords <- pivotCacheRecords[order(nchar(pivotCacheRecords), pivotCacheRecords)]


    wb$pivotDefinitionsRels <- character(nPivotTables)

    pivot_content_type <- NULL

    if (length(pivotTableRelsXML) > 0) {
      wb$pivotTables.xml.rels <- unlist(lapply(pivotTableRelsXML, function(x) removeHeadTag(cppReadFile(x))))
    }


    # ## Check what caches are used
    cache_keep <- unlist(regmatches(wb$pivotTables.xml.rels, gregexpr("(?<=pivotCache/pivotCacheDefinition)[0-9](?=\\.xml)",
      wb$pivotTables.xml.rels,
      perl = TRUE, ignore.case = TRUE
    )))

    ## pivot cache records
    tmp <- unlist(regmatches(pivotCacheRecords, gregexpr("(?<=pivotCache/pivotCacheRecords)[0-9]+(?=\\.xml)", pivotCacheRecords, perl = TRUE, ignore.case = TRUE)))
    pivotCacheRecords <- pivotCacheRecords[tmp %in% cache_keep]

    ## pivot cache definitions rels
    tmp <- unlist(regmatches(pivotDefRelsXML, gregexpr("(?<=_rels/pivotCacheDefinition)[0-9]+(?=\\.xml)", pivotDefRelsXML, perl = TRUE, ignore.case = TRUE)))
    pivotDefRelsXML <- pivotDefRelsXML[tmp %in% cache_keep]

    ## pivot cache definitions
    tmp <- unlist(regmatches(pivotDefXML, gregexpr("(?<=pivotCache/pivotCacheDefinition)[0-9]+(?=\\.xml)", pivotDefXML, perl = TRUE, ignore.case = TRUE)))
    pivotDefXML <- pivotDefXML[tmp %in% cache_keep]



    if (length(pivotTableXML) > 0) {
      wb$pivotTables[seq_along(pivotTableXML)] <- pivotTableXML
      pivot_content_type <- c(
        pivot_content_type,
        sprintf('<Override PartName="/xl/pivotTables/pivotTable%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>', seq_along(pivotTableXML))
      )
    }

    if (length(pivotDefXML) > 0) {
      wb$pivotDefinitions[seq_along(pivotDefXML)] <- pivotDefXML
      pivot_content_type <- c(
        pivot_content_type,
        sprintf('<Override PartName="/xl/pivotCache/pivotCacheDefinition%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>', seq_along(pivotDefXML))
      )
    }

    if (length(pivotCacheRecords) > 0) {
      wb$pivotRecords[seq_along(pivotCacheRecords)] <- pivotCacheRecords
      pivot_content_type <- c(
        pivot_content_type,
        sprintf('<Override PartName="/xl/pivotCache/pivotCacheRecords%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>', seq_along(pivotCacheRecords))
      )
    }

    if (length(pivotDefRelsXML) > 0) {
      wb$pivotDefinitionsRels[seq_along(pivotDefRelsXML)] <- pivotDefRelsXML
    }




    ## update content_types
    wb$Content_Types <- c(wb$Content_Types, pivot_content_type)


    ## workbook rels
    wb$workbook.xml.rels <- c(
      wb$workbook.xml.rels,
      sprintf('<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition%s.xml"/>', rIds, seq_along(pivotDefXML))
    )


    caches <- getNodes(xml = workbook, tagIn = "<pivotCaches>")
    caches <- getChildlessNode(xml = caches, tag = "pivotCache")
    for (i in seq_along(caches)) {
      caches[i] <- gsub('"rId[0-9]+"', sprintf('"rId%s"', rIds[i]), caches[i])
    }

    wb$workbook$pivotCaches <- paste0("<pivotCaches>", paste(caches, collapse = ""), "</pivotCaches>")
  }

  ## xl\vbaProject
  if (length(vbaProject) > 0) {
    wb$vbaProject <- vbaProject
    wb$Content_Types[grepl('<Override PartName="/xl/workbook.xml" ', wb$Content_Types)] <- '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>'
    wb$Content_Types <- c(wb$Content_Types, '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>')
  }


  ## xl\styles
  if (length(stylesXML) > 0) {
    styleObjects <- wb$loadStyles(stylesXML)
  } else {
    styleObjects <- list()
  }

  ## xl\media
  if (length(media) > 0) {
    mediaNames <- regmatches(media, regexpr("image[0-9]+\\.[a-z]+$", media))
    fileTypes <- unique(gsub("image[0-9]+\\.", "", mediaNames))

    contentNodes <- sprintf('<Default Extension="%s" ContentType="image/%s"/>', fileTypes, fileTypes)
    contentNodes[fileTypes == "emf"] <- '<Default Extension="emf" ContentType="image/x-emf"/>'

    wb$Content_Types <- c(contentNodes, wb$Content_Types)
    names(media) <- mediaNames
    wb$media <- media
  }



  ## xl\chart
  if (length(charts) > 0) {
    chartNames <- basename(charts)
    nCharts <- sum(grepl("chart[0-9]+.xml", chartNames))
    nChartStyles <- sum(grepl("style[0-9]+.xml", chartNames))
    nChartCol <- sum(grepl("colors[0-9]+.xml", chartNames))

    if (nCharts > 0) {
      wb$Content_Types <- c(wb$Content_Types, sprintf('<Override PartName="/xl/charts/chart%s.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>', seq_len(nCharts)))
    }

    if (nChartStyles > 0) {
      wb$Content_Types <- c(wb$Content_Types, sprintf('<Override PartName="/xl/charts/style%s.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>', seq_len(nChartStyles)))
    }

    if (nChartCol > 0) {
      wb$Content_Types <- c(wb$Content_Types, sprintf('<Override PartName="/xl/charts/colors%s.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>', seq_len(nChartCol)))
    }

    if (length(chartsRels)) {
      charts <- c(charts, chartsRels)
      chartNames <- c(chartNames, file.path("_rels", basename(chartsRels)))
    }

    names(charts) <- chartNames
    wb$charts <- charts
  }






  ## xl\theme
  if (length(themeXML) > 0) {
    wb$theme <- removeHeadTag(paste(unlist(lapply(sort(themeXML)[[1]], readUTF8)), collapse = ""))
  }


  ## externalLinks
  if (length(extLinksXML) > 0) {
    wb$externalLinks <- lapply(sort(extLinksXML), function(x) removeHeadTag(cppReadFile(x)))

    wb$Content_Types <- c(
      wb$Content_Types,
      sprintf('<Override PartName="/xl/externalLinks/externalLink%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>', seq_along(extLinksXML))
    )

    wb$workbook.xml.rels <- c(wb$workbook.xml.rels, sprintf(
      '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink%s.xml"/>',
      seq_along(extLinksXML)
    ))
  }

  ## externalLinksRels
  if (length(extLinksRelsXML) > 0) {
    wb$externalLinksRels <- lapply(sort(extLinksRelsXML), function(x) removeHeadTag(cppReadFile(x)))
  }







  ##* ----------------------------------------------------------------------------------------------*##
  ### BEGIN READING IN WORKSHEET DATA
  ##* ----------------------------------------------------------------------------------------------*##

  ## xl\worksheets
  file_names <- regmatches(worksheet_rId_mapping, regexpr("sheet[0-9]+\\.xml", worksheet_rId_mapping, perl = TRUE))
  file_rIds <- unlist(getId(worksheet_rId_mapping))
  file_names <- file_names[match(sheetrId, file_rIds)]

  worksheetsXML <- file.path(dirname(worksheetsXML), file_names)
  wb <- loadworksheets(wb = wb, styleObjects = styleObjects, xmlFiles = worksheetsXML, is_chart_sheet = is_chart_sheet)

  ## Fix styleobject encoding
  if (length(wb$styleObjects) > 0) {
    style_names <- sapply(wb$styleObjects, "[[", "sheet")
    Encoding(style_names) <- "UTF-8"
    wb$styleObjects <- lapply(seq_along(style_names), function(i) {
      wb$styleObjects[[i]]$sheet <- style_names[[i]]
      wb$styleObjects[[i]]
    })
  }


  ## Fix headers/footers
  for (i in seq_along(worksheetsXML)) {
    if (!is_chart_sheet[i]) {
      if (length(wb$worksheets[[i]]$headerFooter) > 0) {
        wb$worksheets[[i]]$headerFooter <- lapply(wb$worksheets[[i]]$headerFooter, splitHeaderFooter)
      }
    }
  }


  ##* ----------------------------------------------------------------------------------------------*##
  ### READING IN WORKSHEET DATA COMPLETE
  ##* ----------------------------------------------------------------------------------------------*##


  ## Next sheetRels to see which drawings_rels belongs to which sheet
  if (length(sheetRelsXML) > 0) {

    ## sheetrId is order sheet appears in xlsx file
    ## create a 1-1 vector of rels to worksheet
    ## haveRels is boolean vector where i-the element is TRUE/FALSE if sheet has a rels sheet

    if (length(chartSheetsXML) == 0) {
      allRels <- file.path(dirname(sheetRelsXML[1]), paste0(file_names, ".rels"))
      haveRels <- allRels %in% sheetRelsXML
    } else {
      haveRels <- rep(FALSE, length(wb$worksheets))
      allRels <- rep("", length(wb$worksheets))

      for (i in 1:nSheets) {
        if (is_chart_sheet[i]) {
          ind <- which(chartSheetRIds == sheetrId[i])
          rels_file <- file.path(chartSheetsRelsDir, paste0(chartsheet_rId_mapping[ind], ".rels"))
        } else {
          ind <- sheetrId[i]
          rels_file <- file.path(xmlDir, "xl", "worksheets", "_rels", paste0(file_names[i], ".rels"))
        }
        if (file.exists(rels_file)) {
          allRels[i] <- rels_file
          haveRels[i] <- TRUE
        }
      }
    }

    ## sheet.xml have been reordered to be in the order of sheetrId
    ## not every sheet has a worksheet rels

    xml <- lapply(seq_along(allRels), function(i) {
      if (haveRels[i]) {
        xml <- readUTF8(allRels[[i]])
        xml <- removeHeadTag(xml)
        xml <- gsub("<Relationships .*?>", "", xml)
        xml <- gsub("</Relationships>", "", xml)
        xml <- getChildlessNode(xml = xml, tag = "Relationship")
      } else {
        xml <- "<Relationship >"
      }
      return(xml)
    })




    
    ## Slicers -------------------------------------------------------------------------------------

    

    if (length(slicerXML) > 0) {
      slicerXML <- slicerXML[order(nchar(slicerXML), slicerXML)]
      slicersFiles <- lapply(xml, function(x) as.integer(regmatches(x, regexpr("(?<=slicer)[0-9]+(?=\\.xml)", x, perl = TRUE))))
      inds <- sapply(slicersFiles, length) > 0


      ## worksheet_rels Id for slicer will be rId0
      k <- 1L
      wb$slicers <- rep("", nSheets)
      for (i in 1:nSheets) {

        ## read in slicer[j].XML sheets into sheet[i]
        if (inds[i]) {
          wb$slicers[[i]] <- slicerXML[k]
          k <- k + 1L

          wb$worksheets_rels[[i]] <- unlist(c(
            wb$worksheets_rels[[i]],
            sprintf('<Relationship Id="rId0" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer%s.xml"/>', i)
          ))
          wb$Content_Types <- c(
            wb$Content_Types,
            sprintf('<Override PartName="/xl/slicers/slicer%s.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>', i)
          )

          slicer_xml_exists <- FALSE
          ## Append slicer to worksheet extLst

          if (length(wb$worksheets[[i]]$extLst) > 0) {
            if (grepl('x14:slicer r:id="rId[0-9]+"', wb$worksheets[[i]]$extLst)) {
              wb$worksheets[[i]]$extLst <- sub('x14:slicer r:id="rId[0-9]+"', 'x14:slicer r:id="rId0"', wb$worksheets[[i]]$extLst)
              slicer_xml_exists <- TRUE
            }
          }

          if (!slicer_xml_exists) {
            wb$worksheets[[i]]$extLst <- c(wb$worksheets[[i]]$extLst, genBaseSlicerXML())
          }
        }
      }
    }


    if (length(slicerCachesXML) > 0) {

      ## ---- slicerCaches
      inds <- seq_along(slicerCachesXML)
      wb$Content_Types <- c(wb$Content_Types, sprintf('<Override PartName="/xl/slicerCaches/slicerCache%s.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>', inds))
      wb$slicerCaches <- sapply(slicerCachesXML[order(nchar(slicerCachesXML), slicerCachesXML)], function(x) removeHeadTag(cppReadFile(x)))
      wb$workbook.xml.rels <- c(wb$workbook.xml.rels, sprintf('<Relationship Id="rId%s" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache%s.xml"/>', 1E5 + inds, inds))
      wb$workbook$extLst <- c(wb$workbook$extLst, genSlicerCachesExtLst(1E5 + inds))
    }


    ## Tables --------------------------------------------------------------------------------------

    

    if (length(tablesXML) > 0) {
      tables <- lapply(xml, function(x) as.integer(regmatches(x, regexpr("(?<=table)[0-9]+(?=\\.xml)", x, perl = TRUE))))
      tableSheets <- unlist(lapply(seq_along(sheetrId), function(i) rep(i, length(tables[[i]]))))

      if (length(unlist(tables)) > 0) {
        ## get the tables that belong to each worksheet and create a worksheets_rels for each
        tCount <- 2L ## table r:Ids start at 3
        for (i in seq_along(tables)) {
          if (length(tables[[i]]) > 0) {
            k <- seq_along(tables[[i]]) + tCount
            wb$worksheets_rels[[i]] <- unlist(c(
              wb$worksheets_rels[[i]],
              sprintf('<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table%s.xml"/>', k, k)
            ))


            wb$worksheets[[i]]$tableParts <- sprintf("<tablePart r:id=\"rId%s\"/>", k)
            tCount <- tCount + length(k)
          }
        }

        ## sort the tables into the order they appear in the xml and tables variables
        names(tablesXML) <- basename(tablesXML)
        tablesXML <- tablesXML[sprintf("table%s.xml", unlist(tables))]

        ## tables are now in correct order so we can read them in as they are
        wb$tables <- sapply(tablesXML, function(x) removeHeadTag(paste(readUTF8(x), collapse = "")))

        ## pull out refs and attach names
        refs <- regmatches(wb$tables, regexpr('(?<=ref=")[0-9A-Z:]+', wb$tables, perl = TRUE))
        names(wb$tables) <- refs

        wb$Content_Types <- c(wb$Content_Types, sprintf('<Override PartName="/xl/tables/table%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>', seq_along(wb$tables) + 2))

        ## relabel ids
        for (i in seq_along(wb$tables)) {
          newId <- sprintf(' id="%s" ', i + 2)
          wb$tables[[i]] <- sub(' id="[0-9]+" ', newId, wb$tables[[i]])
        }

        displayNames <- unlist(regmatches(wb$tables, regexpr('(?<=displayName=").*?[^"]+', wb$tables, perl = TRUE)))
        if (length(displayNames) != length(tablesXML)) {
          displayNames <- paste0("Table", seq_along(tablesXML))
        }

        attr(wb$tables, "sheet") <- tableSheets
        attr(wb$tables, "tableName") <- displayNames

        for (i in seq_along(tableSheets)) {
          table_sheet_i <- tableSheets[i]
          attr(wb$worksheets[[table_sheet_i]]$tableParts, "tableName") <- c(attr(wb$worksheets[[table_sheet_i]]$tableParts, "tableName"), displayNames[i])
        }
      }
    } ## if(length(tablesXML) > 0)

    ## might we have some external hyperlinks
    if (any(sapply(wb$worksheets[!is_chart_sheet], function(x) length(x$hyperlinks) > 0))) {

      ## Do we have external hyperlinks
      hlinks <- lapply(xml, function(x) x[grepl("hyperlink", x) & grepl("External", x)])
      hlinksInds <- which(sapply(hlinks, length) > 0)

      ## If it's an external hyperlink it will have a target in the sheet_rels
      if (length(hlinksInds) > 0) {
        for (i in hlinksInds) {
          ids <- unlist(lapply(hlinks[[i]], function(x) regmatches(x, gregexpr('(?<=Id=").*?"', x, perl = TRUE))[[1]]))
          ids <- gsub('"$', "", ids)

          targets <- unlist(lapply(hlinks[[i]], function(x) regmatches(x, gregexpr('(?<=Target=").*?"', x, perl = TRUE))[[1]]))
          targets <- gsub('"$', "", targets)

          ids2 <- lapply(wb$worksheets[[i]]$hyperlinks, function(x) regmatches(x, gregexpr('(?<=r:id=").*?"', x, perl = TRUE))[[1]])
          ids2[sapply(ids2, length) == 0] <- NA
          ids2 <- gsub('"$', "", unlist(ids2))

          targets <- targets[match(ids2, ids)]
          names(wb$worksheets[[i]]$hyperlinks) <- targets
        }
      }
    }



    ## Drawings ------------------------------------------------------------------------------------

    

    ## xml is in the order of the sheets, drawIngs is toes to sheet position of hasDrawing
    ## Not every sheet has a drawing.xml


    drawXMLrelationship <- lapply(xml, function(x) grep("drawings/drawing", x, value = TRUE))
    hasDrawing <- sapply(drawXMLrelationship, length) > 0 ## which sheets have a drawing

    if (length(drawingRelsXML) > 0) {
      dRels <- lapply(drawingRelsXML, readUTF8)
      dRels <- unlist(lapply(dRels, removeHeadTag))
      dRels <- gsub("<Relationships .*?>", "", dRels)
      dRels <- gsub("</Relationships>", "", dRels)
    }

    if (length(drawingsXML) > 0) {
      dXML <- lapply(drawingsXML, readUTF8)
      dXML <- unlist(lapply(dXML, removeHeadTag))
      dXML <- gsub("<xdr:wsDr .*?>", "", dXML)
      dXML <- gsub("</xdr:wsDr>", "", dXML)

      #       ptn1 <- "<(mc:AlternateContent|xdr:oneCellAnchor|xdr:twoCellAnchor|xdr:absoluteAnchor)"
      #       ptn2 <- "</(mc:AlternateContent|xdr:oneCellAnchor|xdr:twoCellAnchor|xdr:absoluteAnchor)>"

      ## split at one/two cell Anchor
      # dXML <- regmatches(dXML, gregexpr(paste0(ptn1, ".*?", ptn2), dXML))
    }


    ## loop over all worksheets and assign drawing to sheet
    if (any(hasDrawing)) {
      for (i in seq_along(xml)) {
        if (hasDrawing[i]) {
          target <- unlist(lapply(drawXMLrelationship[[i]], function(x) regmatches(x, gregexpr('(?<=Target=").*?"', x, perl = TRUE))[[1]]))
          target <- basename(gsub('"$', "", target))

          ## sheet_i has which(hasDrawing)[[i]]
          relsInd <- grepl(target, drawingRelsXML)
          if (any(relsInd)) {
            wb$drawings_rels[i] <- dRels[relsInd]
          }

          drawingInd <- grepl(target, drawingsXML)
          if (any(drawingInd)) {
            wb$drawings[i] <- dXML[drawingInd]
          }
        }
      }
    }




    ## VML Drawings --------------------------------------------------------------------------------


    if (length(vmlDrawingXML) > 0) {
      wb$Content_Types <- c(wb$Content_Types, '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>')

      drawXMLrelationship <- lapply(xml, function(x) grep("drawings/vmlDrawing", x, value = TRUE))
      hasDrawing <- sapply(drawXMLrelationship, length) > 0 ## which sheets have a drawing

      ## loop over all worksheets and assign drawing to sheet
      if (any(hasDrawing)) {
        for (i in seq_along(xml)) {
          if (hasDrawing[i]) {
            target <- unlist(lapply(drawXMLrelationship[[i]], function(x) regmatches(x, gregexpr('(?<=Target=").*?"', x, perl = TRUE))[[1]]))
            target <- basename(gsub('"$', "", target))
            ind <- grepl(target, vmlDrawingXML)

            if (any(ind)) {
              txt <- paste(readUTF8(vmlDrawingXML[ind]), collapse = "\n")
              txt <- removeHeadTag(txt)

              i1 <- regexpr("<v:shapetype", txt, fixed = TRUE)
              i2 <- regexpr("</xml>", txt, fixed = TRUE)

              wb$vml[[i]] <- substring(text = txt, first = i1, last = (i2 - 1L))

              relsInd <- grepl(target, vmlDrawingRelsXML)
              if (any(relsInd)) {
                wb$vml_rels[i] <- vmlDrawingRelsXML[relsInd]
              }
            }
          }
        }
      }
    }







    ## vmlDrawing and comments
    if (length(commentsXML) > 0) {
      drawXMLrelationship <- lapply(xml, function(x) grep("drawings/vmlDrawing[0-9]+\\.vml", x, value = TRUE))
      hasDrawing <- sapply(drawXMLrelationship, length) > 0 ## which sheets have a drawing

      commentXMLrelationship <- lapply(xml, function(x) grep("comments[0-9]+\\.xml", x, value = TRUE))
      hasComment <- sapply(commentXMLrelationship, length) > 0 ## which sheets have a comment

      for (i in seq_along(xml)) {
        if (hasComment[i]) {
          target <- unlist(lapply(drawXMLrelationship[[i]], function(x) regmatches(x, gregexpr('(?<=Target=").*?"', x, perl = TRUE))[[1]]))
          target <- basename(gsub('"$', "", target))
          ind <- grepl(target, vmlDrawingXML)

          if (any(ind)) {
            txt <- paste(readUTF8(vmlDrawingXML[ind]), collapse = "\n")
            txt <- removeHeadTag(txt)

            cd <- unique(getNodes(xml = txt, tagIn = "<x:ClientData"))
            cd <- grep('ObjectType="Note"', cd, value = TRUE)
            cd <- paste0(cd, ">")

            ## now loada comment
            target <- unlist(lapply(commentXMLrelationship[[i]], function(x) regmatches(x, gregexpr('(?<=Target=").*?"', x, perl = TRUE))[[1]]))
            target <- basename(gsub('"$', "", target))

            txt <- paste(readUTF8(grep(target, commentsXML, value = TRUE)), collapse = "\n")
            txt <- removeHeadTag(txt)

            authors <- getNodes(xml = txt, tagIn = "<author>")
            authors <- gsub("<author>|</author>", "", authors)

            comments <- getNodes(xml = txt, tagIn = "<commentList>")
            comments <- gsub("<commentList>", "", comments)
            comments <- getNodes(xml = comments, tagIn = "<comment")

            refs <- regmatches(comments, regexpr('(?<=ref=").*?[^"]+', comments, perl = TRUE))

            authorsInds <- as.integer(regmatches(comments, regexpr('(?<=authorId=").*?[^"]+', comments, perl = TRUE))) + 1
            authors <- authors[authorsInds]

            style <- lapply(comments, getNodes, tagIn = "<rPr>")

            comments <- regmatches(comments, 
                                   gregexpr("(?<=<t( |>))[\\s\\S]+?(?=</t>)", comments, perl = TRUE))
            comments <- lapply(comments, function(x) gsub(".*?>", "", x, perl = TRUE))


            wb$comments[[i]] <- lapply(seq_along(comments), function(j) {
              comment_list <- list(
                "ref" = refs[j],
                "author" = authors[j],
                "comment" = comments[[j]],
                "style" = style[[j]],
                "clientData" = cd[[j]]
              )
            })
          }
        }
      }
    }

    ## Threaded comments
    if (length(threadCommentsXML) > 0) {
      threadCommentsXMLrelationship <- lapply(xml, function(x) grep("threadedComment[0-9]+\\.xml", x, value = TRUE))
      hasThreadComments<- sapply(threadCommentsXMLrelationship, length) > 0
      if(any(hasThreadComments)) {
        for (i in seq_along(xml)) {
          if (hasThreadComments[i]) {
            target <- unlist(lapply(threadCommentsXMLrelationship[[i]], function(x) regmatches(x, gregexpr('(?<=Target=").*?"', x, perl = TRUE))[[1]]))
            target <- basename(gsub('"$', "", target))

            wb$threadComments[[i]] <- grep(target, threadCommentsXML, value = TRUE)
            
          }
        }
      }
      wb$Content_Types <- c(
        wb$Content_Types, 
        sprintf('<Override PartName="/xl/threadedComments/%s" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>',
                sapply(threadCommentsXML, basename))
        )
    }
    
    ## Persons (needed for Threaded Comment)
    if(length(personXML) > 0){
      wb$persons <- personXML
      wb$Content_Types <- c(
        wb$Content_Types,
        '<Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/>'
      )
      wb$workbook.xml.rels <- c(
        wb$workbook.xml.rels, 
        '<Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/>')
    }
    

    ## rels image
    drawXMLrelationship <- lapply(xml, function(x) grep("relationships/image", x, value = TRUE))
    hasDrawing <- sapply(drawXMLrelationship, length) > 0 ## which sheets have a drawing
    if (any(hasDrawing)) {
      for (i in seq_along(xml)) {
        if (hasDrawing[i]) {
          image_ids <- unlist(getId(drawXMLrelationship[[i]]))
          new_image_ids <- paste0("rId", seq_along(image_ids) + 70000)
          for (j in seq_along(image_ids)) {
            wb$worksheets[[i]]$oleObjects <- gsub(image_ids[j], new_image_ids[j], wb$worksheets[[i]]$oleObjects, fixed = TRUE)
            wb$worksheets_rels[[i]] <- c(wb$worksheets_rels[[i]], gsub(image_ids[j], new_image_ids[j], drawXMLrelationship[[i]][j], fixed = TRUE))
          }
        }
      }
    }

    ## rels image
    drawXMLrelationship <- lapply(xml, function(x) grep("relationships/package", x, value = TRUE))
    hasDrawing <- sapply(drawXMLrelationship, length) > 0 ## which sheets have a drawing
    if (any(hasDrawing)) {
      for (i in seq_along(xml)) {
        if (hasDrawing[i]) {
          image_ids <- unlist(getId(drawXMLrelationship[[i]]))
          new_image_ids <- paste0("rId", seq_along(image_ids) + 90000)
          for (j in seq_along(image_ids)) {
            wb$worksheets[[i]]$oleObjects <- gsub(image_ids[j], new_image_ids[j], wb$worksheets[[i]]$oleObjects, fixed = TRUE)
            wb$worksheets_rels[[i]] <- c(
              wb$worksheets_rels[[i]],
              sprintf("<Relationship Id=\"%s\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/package\" Target=\"../embeddings/Microsoft_Word_Document1.docx\"/>", new_image_ids[j])
            )
          }
        }
      }
    }



    ## Embedded docx
    if (length(embeddings) > 0) {
      wb$Content_Types <- c(wb$Content_Types, '<Default Extension="docx" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document"/>')
      wb$embeddings <- embeddings
    }



    ## pivot tables
    if (length(pivotTableXML) > 0) {
      # pivotTableJ <- lapply(xml, function(x) as.integer(regmatches(x, regexpr("(?<=pivotTable)[0-9]+(?=\\.xml)", x, perl = TRUE)))) variable not used
      # sheetWithPivot <- which(sapply(pivotTableJ, length) > 0)  variable not used

      pivotRels <- lapply(xml, function(x) {
        y <- grep("pivotTable", x, value = TRUE)
        y[order(nchar(y), y)]
      })
      hasPivot <- sapply(pivotRels, length) > 0

      ## Modify rIds
      for (i in seq_along(pivotRels)) {
        if (hasPivot[i]) {
          for (j in seq_along(pivotRels[[i]])) {
            pivotRels[[i]][j] <- gsub('"rId[0-9]+"', sprintf('"rId%s"', 20000L + j), pivotRels[[i]][j])
          }

          wb$worksheets_rels[[i]] <- c(wb$worksheets_rels[[i]], pivotRels[[i]])
        }
      }


      ## remove any workbook_res references to pivot tables that are not being used in worksheet_rels
      inds <- seq_along(wb$pivotTables.xml.rels)
      fileNo <- as.integer(unlist(regmatches(unlist(wb$worksheets_rels), gregexpr("(?<=pivotTable)[0-9]+(?=\\.xml)", unlist(wb$worksheets_rels), perl = TRUE))))
      inds <- inds[!inds %in% fileNo]

      if (length(inds) > 0) {
        toRemove <- paste(sprintf("(pivotCacheDefinition%s\\.xml)", inds), collapse = "|")
        fileNo <- grep(toRemove, wb$pivotTables.xml.rels)
        toRemove <- paste(sprintf("(pivotCacheDefinition%s\\.xml)", fileNo), collapse = "|")

        ## remove reference to file from workbook.xml.res
        wb$workbook.xml.rels <- wb$workbook.xml.rels[!grepl(toRemove, wb$workbook.xml.rels)]
      }
    }
  } ## end of worksheetRels

  ## convert hyperliks to hyperlink objects
  for (i in 1:nSheets) {
    wb$worksheets[[i]]$hyperlinks <- xml_to_hyperlink(wb$worksheets[[i]]$hyperlinks)
  }



  ## queryTables
  if (length(queryTablesXML) > 0) {
    ids <- as.numeric(regmatches(queryTablesXML, regexpr("[0-9]+(?=\\.xml)", queryTablesXML, perl = TRUE)))
    wb$queryTables <- unlist(lapply(queryTablesXML[order(ids)], function(x) removeHeadTag(cppReadFile(xmlFile = x))))
    wb$Content_Types <- c(
      wb$Content_Types,
      sprintf('<Override PartName="/xl/queryTables/queryTable%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/>', seq_along(queryTablesXML))
    )
  }


  ## connections
  if (length(connectionsXML) > 0) {
    wb$connections <- removeHeadTag(cppReadFile(xmlFile = connectionsXML))
    wb$workbook.xml.rels <- c(wb$workbook.xml.rels, '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/>')
    wb$Content_Types <- c(wb$Content_Types, '<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>')
  }




  ## table rels
  if (length(tableRelsXML) > 0) {

    ## table_i_might have tableRels_i but I am re-ordering the tables to be in order of worksheets
    ## I make every table have a table_rels so i need to fill in the gaps if any table_rels are missing

    tmp <- paste0(basename(tablesXML), ".rels")
    hasRels <- tmp %in% basename(tableRelsXML)

    ## order tableRelsXML
    tableRelsXML <- tableRelsXML[match(tmp[hasRels], basename(tableRelsXML))]

    ##
    wb$tables.xml.rels <- character(length = length(tablesXML))

    ## which sheet does it belong to
    xml <- sapply(tableRelsXML, cppReadFile, USE.NAMES = FALSE)
    xml <- sapply(xml, removeHeadTag, USE.NAMES = FALSE)

    wb$tables.xml.rels[hasRels] <- xml
  } else if (length(tablesXML) > 0) {
    wb$tables.xml.rels <- rep("", length(tablesXML))
  }


  activesheet <- unlist(regmatches(workbook, gregexpr("(?<=<bookViews>).*(?=</bookViews>)", workbook, perl = TRUE)))
  activesheet <- unlist(regmatches(activesheet, gregexpr("<workbookView[^>]*>", activesheet, perl = TRUE)))
  
  wb$ActiveSheet <- as.integer(getAttrs(activesheet,"activeTab")$activeTab) + 1L
  
  if(length(wb$ActiveSheet) == 0){
    wb$ActiveSheet <- 1L
  }

  return(wb)
}
