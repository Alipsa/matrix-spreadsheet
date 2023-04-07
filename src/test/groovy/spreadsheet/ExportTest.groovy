package spreadsheet

import org.junit.jupiter.api.BeforeAll
import org.junit.jupiter.api.Test
import se.alipsa.groovy.matrix.ConversionException
import se.alipsa.groovy.matrix.TableMatrix
import se.alipsa.groovy.spreadsheet.ExcelExporter
import se.alipsa.groovy.spreadsheet.ExcelReader
import se.alipsa.groovy.spreadsheet.OdsReader
import se.alipsa.groovy.spreadsheet.SpreadSheetExporter
import se.alipsa.groovy.spreadsheet.SpreadsheetUtil

import java.time.LocalDate
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter

import static org.junit.jupiter.api.Assertions.*
import static se.alipsa.groovy.matrix.ListConverter.*

class ExportTest {

  static TableMatrix table
  static TableMatrix table2

  @BeforeAll
  static void init() {
    def matrix = [
        id: [null,2,3,4,-5],
        name: ['foo', 'bar', 'baz', 'bla', null],
        start: toLocalDates('2021-01-04', null, '2023-03-13', '2024-04-15', '2025-05-20'),
        end: toLocalDateTimes(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"), '2021-02-04 12:01:22', '2022-03-12 13:14:15', '2023-04-13 15:16:17', null, '2025-06-20 17:18:19'),
        measure: [12.45, null, 14.11, 15.23, 10.99],
        active: [true, false, null, true, false]
    ]
    table = TableMatrix.create(matrix, [int, String, LocalDate, LocalDateTime, BigDecimal, Boolean])
    def stats = [
        id: [null,2,3,4,-5],
        jan: [1123.1234, 2341.234, 1010.00122, 991, 1100.1],
        feb: [1111.1235, 2312.235, 1001.00121, 999, 1200.7]
    ]
    table2 = TableMatrix.create(stats, [int, BigDecimal, BigDecimal])
  }

  @Test
  void exportExcelTest() {

    //println(table.content())
    def file = File.createTempFile("matrix", ".xlsx")
    SpreadSheetExporter.exportSpreadSheet(file, table)
    println("Wrote to $file")

    SpreadSheetExporter.exportSpreadSheet(file, table2)
    println("Wrote another sheet to $file")
    try ( def reader = new ExcelReader(file)) {
      assertEquals(2, reader.sheetNames.size(), "number of sheets")
    }
  }

  @Test
  void testOdsExport() {
    def odsFile = File.createTempFile("matrix", ".ods")
    if (odsFile.exists()) {
      odsFile.delete()
    }
    SpreadSheetExporter.exportSpreadSheet(odsFile, table, "Sheet 1")
    println("Wrote to $odsFile")
    try ( def reader = new OdsReader(odsFile)) {
      assertEquals(1, reader.sheetNames.size(), "number of sheets")
    }
    SpreadSheetExporter.exportSpreadSheet(odsFile, table2, "Sheet 2")
    println("Wrote another sheet to $odsFile")
    try ( def reader = new OdsReader(odsFile)) {
      assertEquals(2, reader.sheetNames.size(), "number of sheets")
    }
  }

  @Test
  void testValidSheetNames() {
    assertEquals("abl rac adabra ", SpreadsheetUtil.createValidSheetName("abl\\rac[adabra]"))
    assertEquals(" Det var en gång ", SpreadsheetUtil.createValidSheetName("'Det var en gång'"))
    assertEquals("Det var en gång", SpreadsheetUtil.createValidSheetName("Det var en gång"))
  }
}
