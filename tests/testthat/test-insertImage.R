
context("Inserting images")

test_that("Inserting hyperlinked image", {
  # Create a workbook.
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  
  # Insert multiple images. Specifically one with a link before and after other 
  # images to test whether drawings reference the correct index of the drawing 
  # relationships.
  img <- system.file("extdata", "einstein.jpg", package = "openxlsx")
  insertImage(wb, "Sheet 1", img)
  insertImage(wb, "Sheet 1", img, address = "https://github.com/ycphs/openxlsx")
  insertImage(wb, "Sheet 1", img)
  
  # Declare expectations for drawings and drawings_rels xml.
  expected_drawings <- list(
    "<xdr:oneCellAnchor xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"><xdr:from xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">\n    <xdr:col xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">0</xdr:col>\n    <xdr:colOff xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">0</xdr:colOff>\n    <xdr:row xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">0</xdr:row>\n    <xdr:rowOff xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">0</xdr:rowOff>\n  </xdr:from><xdr:ext xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" cx=\"5486400\" cy=\"2743200\"/><xdr:pic xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">\n      <xdr:nvPicPr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">\n        <xdr:cNvPr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" id=\"1\" name=\"Picture 1\"/>\n        <xdr:cNvPicPr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">\n          <a:picLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/>\n        </xdr:cNvPicPr>\n      </xdr:nvPicPr>\n      <xdr:blipFill xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">\n        <a:blip xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rId1\">\n        </a:blip>\n        <a:stretch xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\n          <a:fillRect/>\n        </a:stretch>\n      </xdr:blipFill>\n      <xdr:spPr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">\n        <a:prstGeom xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" prst=\"rect\">\n          <a:avLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>\n        </a:prstGeom>\n      </xdr:spPr>\n    </xdr:pic><xdr:clientData/></xdr:oneCellAnchor>",
    "<xdr:oneCellAnchor xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"><xdr:from xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">\n    <xdr:col xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">0</xdr:col>\n    <xdr:colOff xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">0</xdr:colOff>\n    <xdr:row xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">0</xdr:row>\n    <xdr:rowOff xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">0</xdr:rowOff>\n  </xdr:from><xdr:ext xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" cx=\"5486400\" cy=\"2743200\"/><xdr:pic xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">\n      <xdr:nvPicPr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">\n        <xdr:cNvPr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" id=\"2\" name=\"Picture 2\">\n              <a:hlinkClick xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId3\"/>\n            </xdr:cNvPr>\n        <xdr:cNvPicPr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">\n          <a:picLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/>\n        </xdr:cNvPicPr>\n      </xdr:nvPicPr>\n      <xdr:blipFill xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">\n        <a:blip xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rId2\">\n        </a:blip>\n        <a:stretch xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\n          <a:fillRect/>\n        </a:stretch>\n      </xdr:blipFill>\n      <xdr:spPr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">\n        <a:prstGeom xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" prst=\"rect\">\n          <a:avLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>\n        </a:prstGeom>\n      </xdr:spPr>\n    </xdr:pic><xdr:clientData/></xdr:oneCellAnchor>",
    "<xdr:oneCellAnchor xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"><xdr:from xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">\n    <xdr:col xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">0</xdr:col>\n    <xdr:colOff xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">0</xdr:colOff>\n    <xdr:row xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">0</xdr:row>\n    <xdr:rowOff xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">0</xdr:rowOff>\n  </xdr:from><xdr:ext xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" cx=\"5486400\" cy=\"2743200\"/><xdr:pic xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">\n      <xdr:nvPicPr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">\n        <xdr:cNvPr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" id=\"3\" name=\"Picture 3\"/>\n        <xdr:cNvPicPr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">\n          <a:picLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/>\n        </xdr:cNvPicPr>\n      </xdr:nvPicPr>\n      <xdr:blipFill xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">\n        <a:blip xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rId4\">\n        </a:blip>\n        <a:stretch xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\n          <a:fillRect/>\n        </a:stretch>\n      </xdr:blipFill>\n      <xdr:spPr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">\n        <a:prstGeom xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" prst=\"rect\">\n          <a:avLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>\n        </a:prstGeom>\n      </xdr:spPr>\n    </xdr:pic><xdr:clientData/></xdr:oneCellAnchor>"
  )
  expected_drawings_rels <- list(
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image1.jpg\"/>",
    "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image2.jpg\"/>",
    "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://github.com/ycphs/openxlsx\" TargetMode=\"External\"/>",
    "<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image3.jpg\"/>"
  )
  
  # Test expectations for drawings and drawings_rels xml.
  expect_equal(expected_drawings, wb$drawings[[1]])
  expect_equal(expected_drawings_rels, wb$drawings_rels[[1]])
  
  # Test errors.
  expect_error(
    insertImage(wb, "Sheet 1", img, address = NULL), 
    "Invalid address"
  )
  expect_error(
    insertImage(wb, "Sheet 1", img, address = NA), 
    "Invalid address"
  )
  expect_error(
    insertImage(
      wb, 
      "Sheet 1", 
      img, 
      address = c("https://github.com/ycphs/openxlsx", "https://github.com/ycphs/openxlsx/issues")
    ), 
    "Invalid address"
  )
})
