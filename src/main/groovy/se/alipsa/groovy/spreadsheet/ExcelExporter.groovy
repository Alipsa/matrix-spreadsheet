package se.alipsa.groovy.spreadsheet

import org.apache.logging.log4j.LogManager
import org.apache.logging.log4j.Logger
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import se.alipsa.groovy.matrix.TableMatrix
import se.alipsa.groovy.matrix.ValueConverter

import java.time.LocalDate
import java.time.LocalDateTime

class ExcelExporter {

  static final Logger logger = LogManager.getLogger()

  static void exportExcel(String filePath, TableMatrix data) {
    File file = new File(filePath)
    exportExcel(file, data)
  }

  static void exportExcel(File file, TableMatrix data) {

    try (Workbook workbook = WorkbookFactory.create(isXssf(file))) {
      Sheet sheet = workbook.createSheet();
      buildSheet(data, sheet)
      writeFile(file, workbook)
    }
  }

  private static boolean isXssf(File file) {
    return !file.getName().toLowerCase().endsWith(".xls");
  }

  private static void buildSheet(TableMatrix data, Sheet sheet) {
    def creationHelper = sheet.getWorkbook().getCreationHelper()
    def localDateStyle = sheet.getWorkbook().createCellStyle()
    localDateStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd"))
    def localDateTimeStyle = sheet.getWorkbook().createCellStyle()
    localDateTimeStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"))

    def names = data.columnNames()
    Row headerRow = sheet.createRow(0)
    for (int i = 0; i < names.size(); i++) {
      headerRow.createCell(i).setCellValue(names[i])
    }

    int rowIdx = 1
    for (matrixRow in data.rows()) {
      //println "MatrixRow = $matrixRow"
      for (int col = 0; col < data.columnCount(); col++) {
        Row row = sheet.getRow(rowIdx)
        if (row == null) row = sheet.createRow(rowIdx)
        Cell cell = row.createCell(col)
        Class type = data.columnType(col)

        if (type in [double, Double, BigDecimal, float, Float, Long, long, BigInteger, Number]) {
          cell.setCellValue(ValueConverter.toDouble(matrixRow[col]))
        } else if (type in [int, Integer, short, Short]) {
          cell.setCellValue(ValueConverter.toInteger(matrixRow[col]))
        } else if (type in [byte, Byte]) {
          cell.setCellValue(matrixRow[col] as Byte)
        } else if (boolean == type || Boolean == type) {
          cell.setCellValue(ValueConverter.toBoolean(matrixRow[col]))
        } else if (LocalDate == type) {
          cell.setCellValue(ValueConverter.toLocalDate(matrixRow[col]))
          cell.setCellStyle(localDateStyle)
        } else if (LocalDateTime == type) {
          cell.setCellValue(ValueConverter.toLocalDateTime(matrixRow[col]))
          cell.setCellStyle(localDateTimeStyle)
        } else if (Date == type) {
          cell.setCellValue(matrixRow[col] as Date)
          cell.setCellStyle(localDateTimeStyle)
        } else {
          def val = matrixRow[col]
          if (val == null) {
            cell.setBlank()
          } else {
            cell.setCellValue(String.valueOf(val))
          }
        }
      }
      rowIdx++
    }
  }

  private static void writeFile(File file, Workbook workbook) throws IOException {
    if (workbook == null) {
      logger.warn("Workbook is null, cannot write to file");
      return
    }
    logger.info("Writing spreadsheet to {}", file.getAbsolutePath());
    try(FileOutputStream fos = new FileOutputStream(file)) {
      workbook.write(fos)
    }
  }
}
