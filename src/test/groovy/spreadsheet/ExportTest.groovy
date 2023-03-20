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
        id: [1,2,3,4,5],
        name: ['foo', 'bar', 'baz', 'bla', 'gnu'],
        start: toLocalDates('2021-01-04', '2022-02-12', '2023-03-13', '2024-04-15', '2025-05-20'),
        end: toLocalDateTimes('2021-02-04 12:01:22', '2022-03-12 13:14:15', '2023-04-13 15:16:17', '2024-05-15 16:17:18', '2025-06-20 17:18:19'),
        measure: [12.45, 13.55, 14.11, 15.23, 10.99]
    ]
    def table = TableMatrix.create(matrix, [int, String, LocalDate, LocalDateTime, BigDecimal])
    //println(table.content())
    def file = File.createTempFile("matrix", ".xlsx")
    ExcelExporter.exportExcel(file, table)
    println("Wrote to $file")
    file.deleteOnExit()
  }

  static List<LocalDateTime> toLocalDateTimes(String[] dates) {
    def formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")
    def dat = new ArrayList<LocalDateTime>(dates.length)
    dates.eachWithIndex { String d, int i ->
      try {
        dat.add(LocalDateTime.parse(d, formatter))
      } catch (Exception e) {
        throw new ConversionException("Failed to convert $d to LocalDateTime in index $i", e)
      }
    }
    return dat
  }
}
