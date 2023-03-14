package se.alipsa.groovy.spreadsheet

import se.alipsa.groovy.matrix.*
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory

class ExcelImporter {

    static final String FILE_PATH = 'filePath'
    static final String SHEET_NAME = 'sheetName'
    static final String START_ROW_NUM = 'startRowNum'
    static final String END_ROW_NUM = 'endRowNum'
    static final String START_COL_NUM = 'startColNum'
    static final String END_COL_NUM = 'endColNum'
    static final String FIRST_ROW_AS_COL_NAMES = 'firstRowAsColNames'


    static TableMatrix importExcelSheet(Map params) {
        String filePath = params.getOrDefault(FILE_PATH, null)
        String sheetName = params.getOrDefault(SHEET_NAME, 'Sheet1')
        Integer startRowNum = params.getOrDefault(START_ROW_NUM,1) as Integer
        Integer endRowNum = params.getOrDefault(END_ROW_NUM, null) as Integer
        Integer startColNum = params.getOrDefault(START_COL_NUM, 1) as Integer
        int endColNum = params.getOrDefault(END_COL_NUM, null) as Integer
        boolean firstRowAsColNames = params.getOrDefault(FIRST_ROW_AS_COL_NAMES, false) as boolean
        return importExcelSheet(filePath, sheetName, startRowNum, endRowNum, startColNum, endColNum, firstRowAsColNames)
    }

    static TableMatrix importExcelSheet(String filePath, String sheetName = 'Sheet1',
                                          int startRowNum = 1, int endRowNum,
                                          int startColNum = 1, int endColNum,
                                          boolean firstRowAsColNames = false) {
        def header = []
        File excelFile = FileUtil.checkFilePath(filePath);
        try (Workbook workbook = WorkbookFactory.create(excelFile)) {
            Sheet sheet = workbook.getSheet(sheetName)
            if (firstRowAsColNames) {
                buildHeaderRow(startRowNum, startColNum, endColNum, header, sheet)
                startRowNum = startRowNum + 1
            } else {
                for (int i = 1; i <= endColNum - startColNum; i++) {
                    header.add(String.valueOf(i))
                }
            }
            return importExcel(sheet, startRowNum, endRowNum, startColNum, endColNum, header)
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
}
