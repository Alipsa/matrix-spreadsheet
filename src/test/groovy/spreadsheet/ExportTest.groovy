package spreadsheet

import org.junit.jupiter.api.Test
import se.alipsa.groovy.matrix.ConversionException
import se.alipsa.groovy.matrix.TableMatrix
import se.alipsa.groovy.spreadsheet.ExcelExporter

import java.time.LocalDate
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter

import static org.junit.jupiter.api.Assertions.*
import static se.alipsa.groovy.matrix.ListConverter.*

class ExportTest {

  @Test
  void exportExcelTest() {
    def matrix = [
        id: [null,2,3,4,-5],
        name: ['foo', 'bar', 'baz', 'bla', null],
        start: toLocalDates('2021-01-04', null, '2023-03-13', '2024-04-15', '2025-05-20'),
        end: toLocalDateTimes(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"), '2021-02-04 12:01:22', '2022-03-12 13:14:15', '2023-04-13 15:16:17', null, '2025-06-20 17:18:19'),
        measure: [12.45, null, 14.11, 15.23, 10.99],
        active: [true, false, null, true, false]
    ]
    def table = TableMatrix.create(matrix, [int, String, LocalDate, LocalDateTime, BigDecimal, Boolean])
    //println(table.content())
    def file = File.createTempFile("matrix", ".xlsx")
    ExcelExporter.exportExcel(file, table)
    println("Wrote to $file")
    //file.deleteOnExit()
  }
}
