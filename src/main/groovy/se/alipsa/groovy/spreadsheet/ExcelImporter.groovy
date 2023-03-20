package se.alipsa.groovy.spreadsheet

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DateUtil
import se.alipsa.groovy.matrix.*
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory

class ExcelImporter {

    static TableMatrix importExcelSheet(Map params) {
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

        return importExcelSheet(
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
    static TableMatrix importExcelSheet(String file, String sheetName = 'Sheet1',
                                        int startRow = 1, int endRow,
                                        String startCol = 'A', String endCol,
                                        boolean firstRowAsColNames = true) {

        return importExcelSheet(
                file,
                sheetName,
                startRow as int,
                endRow as int,
                SpreadsheetUtil.asColumnNumber(startCol) as int,
                SpreadsheetUtil.asColumnNumber(endCol) as int,
                firstRowAsColNames as boolean
        )
    }

    static TableMatrix importExcelSheet(String file, String sheetName = 'Sheet1',
                                          int startRow = 1, int endRow,
                                          int startCol = 1, int endCol,
                                          boolean firstRowAsColNames = true) {
        def header = []
        File excelFile = FileUtil.checkFilePath(file);
        try (Workbook workbook = WorkbookFactory.create(excelFile)) {
            Sheet sheet = workbook.getSheet(sheetName)
            if (firstRowAsColNames) {
                buildHeaderRow(startRow, startCol, endCol, header, sheet)
                startRow = startRow + 1
            } else {
                for (int i = 1; i <= endCol - startCol; i++) {
                    header.add(String.valueOf(i))
                }
            }
            return importExcel(sheet, startRow, endRow, startCol, endCol, header)
        }
    }

    private static void buildHeaderRow(int startRowNum, int startColNum, int endColNum, List<String> header, Sheet sheet) {
        startRowNum--
        startColNum--
        endColNum--
        ExcelValueExtractor ext = new ExcelValueExtractor(sheet)
        Row row = sheet.getRow(startRowNum)
        for (int i = 0; i <= endColNum - startColNum; i++) {
            header.add(ext.getString(row, startColNum + i))
        }
    }

    private static TableMatrix importExcel(Sheet sheet, int startRowNum, int endRowNum, int startColNum, int endColNum, List<String> colNames) {
        startRowNum--
        endRowNum--
        startColNum--
        endColNum--

        ExcelValueExtractor ext = new ExcelValueExtractor(sheet)
        List<List<?>> matrix = []
        List<?> rowList
        for (int rowIdx = startRowNum; rowIdx <= endRowNum; rowIdx++) {
            Row row = sheet.getRow(rowIdx)
            rowList = []
            for (int colIdx = startColNum; colIdx <= endColNum; colIdx++) {
                def cell = row.getCell(colIdx)
                if (cell == null) {
                    rowList.add(null)
                } else {
                    //println("cell[$rowIdx, $colIdx]: ${cell.getCellType()}: val: ${ext.getObject(cell)}")
                    switch (cell.getCellType()) {
                        case CellType.BLANK -> rowList.add(null)
                        case CellType.BOOLEAN -> rowList.add(ext.getBoolean(row, colIdx))
                        case CellType.NUMERIC -> {
                            if (DateUtil.isCellDateFormatted(cell)) {
                                rowList.add(ext.getLocalDateTime(cell))
                            } else {
                                rowList.add(ext.getDouble(row, colIdx))
                            }
                        }
                        case CellType.STRING -> rowList.add(ext.getString(row, colIdx))
                        case CellType.FORMULA -> rowList.add(ext.getObject(cell))
                        default -> rowList.add(ext.getString(row, colIdx))
                    }
                }
            }
            matrix.add(rowList)
        }
        return TableMatrix.create(colNames, matrix, [Object]*colNames.size())
    }

    static void validateNotNull(Object paramVal, String paramName) {
        if (paramVal == null) {
            throw new IllegalArgumentException("$paramName cannot be null")
        }
    }
}
