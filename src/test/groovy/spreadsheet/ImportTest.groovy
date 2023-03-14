package spreadsheet

import org.junit.jupiter.api.Test
import se.alipsa.groovy.spreadsheet.ExcelImporter
import java.time.format.DateTimeFormatter

import java.time.LocalDate

import static org.junit.jupiter.api.Assertions.*

class ImportTest {

    @Test
    void testExcelImport() {
        def table = ExcelImporter.importExcelSheet(filePath: "Book1.xlsx", endRowNum: 11, endColNum: 4, firstRowAsColNames: true)
        table = table.convert(id: Integer, bar: LocalDate, baz: BigDecimal, DateTimeFormatter.ofPattern('yyyy-MM-dd HH:mm:ss.SSS'))
        //println(table.content())
        assertEquals(3, table[2, 0])
        assertEquals(LocalDate.parse("2023-05-06"), table[5, 2])
        assertEquals(17.4, table['baz'][table.rowCount()-1])
        assertEquals(['id', 'foo', 'bar', 'baz'], table.columnNames())
    }
}
