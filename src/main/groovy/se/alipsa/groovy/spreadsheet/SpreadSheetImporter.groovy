package se.alipsa.groovy.spreadsheet

import se.alipsa.groovy.matrix.TableMatrix

class SpreadSheetImporter {

  static TableMatrix importSpreadsheet(String file, String sheetName = 'Sheet1',
                                      int startRow = 1, int endRow,
                                      int startCol = 1, int endCol,
                                      boolean firstRowAsColNames = true) {
    if (file.toLowerCase().endsWith(".ods")) {
      return OdsImporter.importOds(file, sheetName, startRow, endRow, startCol, endCol, firstRowAsColNames)
    }
    return ExcelImporter.importExcelSheet(file, sheetName, startRow, endRow, startCol, endCol, firstRowAsColNames)
  }

  static TableMatrix importSpreadsheet(String file, String sheetName = 'Sheet1',
                                      int startRow = 1, int endRow,
                                      String startCol = 'A', String endCol,
                                      boolean firstRowAsColNames = true) {
    if (file.toLowerCase().endsWith(".ods")) {
      return OdsImporter.importOds(file, sheetName, startRow, endRow, startCol, endCol, firstRowAsColNames)
    }
    return ExcelImporter.importExcelSheet(file, sheetName, startRow, endRow, startCol, endCol, firstRowAsColNames)
  }

  static TableMatrix importSpreadsheet(Map params) {
    def fp = params.getOrDefault('file', null)
    validateNotNull(fp, 'file')
    String file
    if (fp instanceof File) {
      file = fp.getAbsolutePath()
    } else {
      file = String.valueOf(fp)
    }
    if (file.toLowerCase().endsWith(".ods")) {
      return OdsImporter.importOds(params)
    }
    return ExcelImporter.importExcelSheet(params)
  }

  static void validateNotNull(Object paramVal, String paramName) {
    if (paramVal == null) {
      throw new IllegalArgumentException("$paramName cannot be null")
    }
  }
}
