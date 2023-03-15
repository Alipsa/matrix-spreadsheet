package se.alipsa.groovy.spreadsheet

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

        ExcelValueExtractor ext = new ExcelValueExtractor(sheet);
        int numRows = 0
        List<List<?>> matrix = []
        List<?> rowList
        for (int rowIdx = startRowNum; rowIdx <= endRowNum; rowIdx++) {
            numRows++;
            Row row = sheet.getRow(rowIdx)
            rowList = []
            int i = 0
            for (int colIdx = startColNum; colIdx <= endColNum; colIdx++) {
                //System.out.println("Adding ext.getString(" + rowIdx + ", " + colIdx+ ") = " + ext.getString(row, colIdx));
                //builders.get(i++).add(ext.getString(row, colIdx));
                String val = ext.getString(row, colIdx)
                rowList.add(val)
            }
            matrix.add(rowList)
        }
        return TableMatrix.create(colNames, matrix, [String]*colNames.size())
    }

    static void validateNotNull(Object paramVal, String paramName) {
        if (paramVal == null) {
            throw new IllegalArgumentException("$paramName cannot be null")
        }
    }
}
