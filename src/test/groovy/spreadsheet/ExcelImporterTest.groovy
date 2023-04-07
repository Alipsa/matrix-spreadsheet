package spreadsheet

import org.junit.jupiter.api.Test

import java.time.LocalDateTime
import java.time.format.DateTimeFormatter

import java.time.LocalDate

import static se.alipsa.groovy.spreadsheet.ExcelImporter.*
import static org.junit.jupiter.api.Assertions.*

class ExcelImporterTest {

    @Test
    void testExcelImport() {
        def table = importExcelSheet(file: "Book1.xlsx", endRow: 11, endCol: 4, firstRowAsColNames: true)
        table = table.convert(id: Integer, bar: LocalDate, baz: BigDecimal, DateTimeFormatter.ofPattern('yyyy-MM-dd HH:mm:ss.SSS'))
        //println(table.content())
        assertEquals(3, table[2, 0])
        assertEquals(LocalDate.parse("2023-05-06"), table[5, 2])
        assertEquals(17.4, table['baz'][table.rowCount()-1])
        assertEquals(['id', 'foo', 'bar', 'baz'], table.columnNames())
    }

    @Test
    void TestImportWithColnames() {
        def table = importExcelSheet(
                "file": "Book1.xlsx",
                "endRow": 11,
                "startCol": 'A',
                "endCol": 'D'
        )
        //println(table.content())
        assertEquals(3.0d, table[2, 0])
        assertEquals(LocalDateTime.parse("2023-05-06T00:00:00.000"), table[5, 2])
        assertEquals(17.4, table['baz'][table.rowCount()-1])
        assertEquals(['id', 'foo', 'bar', 'baz'], table.columnNames())
    }
}
