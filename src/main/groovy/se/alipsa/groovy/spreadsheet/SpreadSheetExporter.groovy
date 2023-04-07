package se.alipsa.groovy.spreadsheet

import org.apache.poi.ss.util.WorkbookUtil
import se.alipsa.groovy.matrix.TableMatrix

class SpreadSheetExporter {

  static String exportSpreadsheet(File file, TableMatrix data) {
    if (file.getName().toLowerCase().endsWith(".ods")) {
      return OdsExporter.exportOds(file, data)
    }
    return ExcelExporter.exportExcel(file, data)
  }

  static String exportSpreadsheet(File file, TableMatrix data, String sheetName) {
    if (file.getName().toLowerCase().endsWith(".ods")) {
      return OdsExporter.exportOds(file, data, sheetName)
    }
    return ExcelExporter.exportExcel(file, data, sheetName)
  }
}
