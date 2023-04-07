package se.alipsa.groovy.spreadsheet

import com.github.miachm.sods.Sheet
import com.github.miachm.sods.SpreadSheet
import se.alipsa.groovy.matrix.TableMatrix
import static se.alipsa.groovy.spreadsheet.SpreadSheetImporter.validateNotNull

/**
 * Import Calc (ods file) into Renjin R in the form of a data.frame (ListVector).
 */
class OdsImporter {

  static TableMatrix importOds(Map params) {
    def fp = params.getOrDefault('file', null)
    String file
    if (fp instanceof File) {
      file = fp.getAbsolutePath()
    } else {
      file = String.valueOf(fp)
    }
    validateNotNull(file, 'file')
    String sheetName = params.getOrDefault('sheetName', 'Sheet1')
    validateNotNull(sheetName, 'sheetName')
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

    return importOds(
        file,
        sheetName,
        startRow as int,
        endRow as int,
        startCol as int,
        endCol as int,
        firstRowAsColNames as boolean
    )
  }

  /**
   * Import an excel spreadsheet
   * @param file the filePath or the file object pointing to the excel file
   * @param sheetName the name of the sheet to import, default is 'Sheet1'
   * @param startRow the starting row for the import (as you would see the row number in excel), defaults to 1
   * @param endRow the last row to import
   * @param startCol the starting column name (A, B etc) or column number (1, 2 etc.)
   * @param endCol the end column name (K, L etc) or column number (11, 12 etc.)
   * @param firstRowAsColNames whether the first row should be used for the names of each column, if false
   * it column names will be v1, v2 etc. Defaults to true
   * @return A TableMatrix with the excel data.
   */
  static TableMatrix importOds(String file, String sheetName = 'Sheet1',
                                      int startRow = 1, int endRow,
                                      String startCol = 'A', String endCol,
                                      boolean firstRowAsColNames = true) {

    return importOds(
        file,
        sheetName,
        startRow as int,
        endRow as int,
        SpreadsheetUtil.asColumnNumber(startCol) as int,
        SpreadsheetUtil.asColumnNumber(endCol) as int,
        firstRowAsColNames as boolean
    )
  }

  static TableMatrix importOds(String file, String sheetName = 'Sheet1',
                                      int startRow = 1, int endRow,
                                      int startCol = 1, int endCol,
                                      boolean firstRowAsColNames = true) {
    def header = []
    File excelFile = FileUtil.checkFilePath(file)
    SpreadSheet spreadSheet = new SpreadSheet(excelFile)
    Sheet sheet = spreadSheet.getSheet(sheetName)
    if (firstRowAsColNames) {
      buildHeaderRow(startRow, startCol, endCol, header, sheet)
      startRow = startRow + 1
    } else {
      for (int i = 1; i <= endCol - startCol; i++) {
        header.add(String.valueOf(i))
      }
    }
    return importOds(sheet, startRow, endRow, startCol, endCol, header)
  }

  static TableMatrix importOds(Sheet sheet, int startRow, int endRow, int startCol, int endCol, List<String> colNames) {
    startRow--
    endRow--
    startCol--
    endCol--

    OdsValueExtractor ext = new OdsValueExtractor(sheet)
    List<List<?>> matrix = []
    List<?> rowList
    for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
      rowList = []
      for (int colIdx = startCol; colIdx <= endCol; colIdx++) {
        String val = ext.getString(rowIdx, colIdx)
        if (val == null) {
          rowList.add(null)
        } else {
          if (val.endsWith("%")) {
            // In excel there is no % in the end of percentage cell, so we make it the same
            // This has the unfortunate consequence that intentional columns ending with % will be changed
            try {
              double dblVal = Double.parseDouble(val.replace("%", "").replace(",", ".")) / 100;
              val = String.valueOf(dblVal)
            } catch (NumberFormatException ignored) {
              // it is not a percentage number, leave the value as it was
            }
          }
          rowList.add(val)
        }
      }
      matrix.add(rowList)
    }
    return TableMatrix.create(colNames, matrix, [String]*colNames.size())
  }

  private static void buildHeaderRow(int startRowNum, int startColNum, int endColNum, List<String> header, Sheet sheet) {
    startRowNum--
    startColNum--
    endColNum--
    OdsValueExtractor ext = new OdsValueExtractor(sheet)
    for (int i = 0; i <= endColNum - startColNum; i++) {
      header.add(ext.getString(startRowNum, startColNum + i))
    }
  }
}
