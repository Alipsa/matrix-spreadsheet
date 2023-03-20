# spreadsheet
Groovy spreadsheet import/export

Here is a simple example:

## Importing a spreadsheet
```groovy
import static se.alipsa.groovy.spreadsheet.ExcelImporter.*
import se.alipsa.groovy.matrix.TableMatrix

TableMatrix table = importExcelSheet(file: "Book1.xlsx", endRow: 11, endCol: 4)
```
The ExcelImporter.importExcelSheet takes the following parameters:
- _file_ the filePath or the file object pointing to the Excel file
- _sheetName_ the name of the sheet to import, default is 'Sheet1'
- _startRow_ the starting row for the import (as you would see the row number in Excel), defaults to 1
- _endRow_ the last row to import
- _startCol_ the starting column name (A, B etc.) or column number (1, 2 etc.)
- _endCol_ the end column name (K, L etc) or column number (11, 12 etc.)
- _firstRowAsColNames_ whether the first row should be used for the names of each column, if false the column names will be v1, v2 etc. Defaults to true

## Export a spreadsheet

```groovy
import static se.alipsa.groovy.spreadsheet.ExcelExporter
import se.alipsa.groovy.matrix.TableMatrix

def table = TableMatrix.create(
    [
        id: [null,2,3,4,-5],
        name: ['foo', 'bar', 'baz', 'bla', null],
        start: toLocalDates('2021-01-04', null, '2023-03-13', '2024-04-15', '2025-05-20'),
        end: toLocalDateTimes(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"), '2021-02-04 12:01:22', '2022-03-12 13:14:15', '2023-04-13 15:16:17', null, '2025-06-20 17:18:19'),
        measure: [12.45, null, 14.11, 15.23, 10.99],
        active: [true, false, null, true, false]
    ]
    , [int, String, LocalDate, LocalDateTime, BigDecimal, Boolean]
)
def file = File.createTempFile("matrix", ".xlsx")

// Export the TableMatrix to an excel file
ExcelExporter.exportExcel(file, table)
```