
openxlsxFontSizeLookupTable <- utils::read.csv("data-raw/openxlsxFontSizeLookupTable.csv")
openxlsxFontSizeLookupTableBold <- utils::read.csv("data-raw/openxlsxFontSizeLookupTableBold.csv")

usethis::use_data(
  openxlsxFontSizeLookupTable,
  openxlsxFontSizeLookupTableBold,
  internal = TRUE
)

# usethis::use_data(openxlsxFontSizeLookupTable, overwrite = TRUE)
# usethis::use_data(openxlsxFontSizeLookupTableBold, overwrite = TRUE)
