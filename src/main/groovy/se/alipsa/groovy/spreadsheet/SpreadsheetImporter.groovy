package se.alipsa.groovy.spreadsheet

import se.alipsa.groovy.matrix.TableMatrix
import se.alipsa.groovy.spreadsheet.excel.ExcelImporter
import se.alipsa.groovy.spreadsheet.ods.OdsImporter

class SpreadsheetImporter {

  static TableMatrix importSpreadsheet(String file, int sheet,
                                       int startRow = 1, int endRow,
                                       int startColumn = 1, int endColumn,
                                       boolean firstRowAsColNames = true) {
    if (file.toLowerCase().endsWith(".ods")) {
      return OdsImporter.importOds(file, sheet, startRow, endRow, startColumn, endColumn, firstRowAsColNames)
    }
    return ExcelImporter.importExcel(file, sheet, startRow, endRow, startColumn, endColumn, firstRowAsColNames)
  }

  static TableMatrix importSpreadsheet(String file, int sheet,
                                       int startRow = 1, int endRow,
                                       String startColumn = 'A', String endColumn,
                                       boolean firstRowAsColNames = true) {
    if (file.toLowerCase().endsWith(".ods")) {
      return OdsImporter.importOds(file, sheet, startRow, endRow, startColumn, endColumn, firstRowAsColNames)
    }
    return ExcelImporter.importExcel(file, sheet, startRow, endRow, startColumn, endColumn, firstRowAsColNames)
  }

  static TableMatrix importSpreadsheet(String file, String sheet = 'Sheet1',
                                      int startRow = 1, int endRow,
                                      int startCol = 1, int endCol,
                                      boolean firstRowAsColNames = true) {
    if (file.toLowerCase().endsWith(".ods")) {
      return OdsImporter.importOds(file, sheet, startRow, endRow, startCol, endCol, firstRowAsColNames)
    }
    return ExcelImporter.importExcel(file, sheet, startRow, endRow, startCol, endCol, firstRowAsColNames)
  }

  static TableMatrix importSpreadsheet(String file, String sheet = 'Sheet1',
                                      int startRow = 1, int endRow,
                                      String startCol = 'A', String endCol,
                                      boolean firstRowAsColNames = true) {
    if (file.toLowerCase().endsWith(".ods")) {
      return OdsImporter.importOds(file, sheet, startRow, endRow, startCol, endCol, firstRowAsColNames)
    }
    return ExcelImporter.importExcel(file, sheet, startRow, endRow, startCol, endCol, firstRowAsColNames)
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
    def sheet = params.get("sheet")
    if (sheet == null) {
      sheet = params.get("sheetNumber")
      if (sheet == null) {
        sheet = params.getOrDefault('sheetName', 'Sheet1')
      }
    }
    validateNotNull(sheet, 'sheet')
    Integer startRow = params.getOrDefault('startRow',1) as Integer
    validateNotNull(startRow, 'startRow')
    Integer endRow = params.getOrDefault('endRow', null) as Integer
    validateNotNull(endRow, 'endRow')
    def startCol = params.getOrDefault('startCol', 1)
    if (startCol instanceof String) {
      startCol = SpreadsheetUtil.asColumnNumber(startCol)
    }
    validateNotNull(startCol, 'startCol')
    def endCol = params.get('endCol')
    if (endCol instanceof String) {
      endCol = SpreadsheetUtil.asColumnNumber(endCol)
    }
    validateNotNull(endCol, 'endCol')
    Boolean firstRowAsColNames = params.getOrDefault('firstRowAsColNames', true) as Boolean
    validateNotNull(firstRowAsColNames, 'firstRowAsColNames')

    if (file.toLowerCase().endsWith(".ods")) {
      if (sheet instanceof Integer) {
        return OdsImporter.importOds(
            file,
            sheet as int,
            startRow as int,
            endRow as int,
            startCol as int,
            endCol as int,
            firstRowAsColNames as boolean
        )
      }
      return OdsImporter.importOds(
          file,
          sheet as String,
          startRow as int,
          endRow as int,
          startCol as int,
          endCol as int,
          firstRowAsColNames as boolean
      )
    }
    if (sheet instanceof Integer) {
      return ExcelImporter.importExcel(
          file,
          sheet as int,
          startRow as int,
          endRow as int,
          startCol as int,
          endCol as int,
          firstRowAsColNames as boolean
      )
    }
    return ExcelImporter.importExcel(
        file,
        sheet as String,
        startRow as int,
        endRow as int,
        startCol as int,
        endCol as int,
        firstRowAsColNames as boolean
    )
  }

  static void validateNotNull(Object paramVal, String paramName) {
    if (paramVal == null) {
      throw new IllegalArgumentException("$paramName cannot be null")
    }
  }
}
