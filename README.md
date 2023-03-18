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
- _file_ the filePath or the file object pointing to the excel file
- _sheetName_ the name of the sheet to import, default is 'Sheet1'
- _startRow_ the starting row for the import (as you would see the row number in excel), defaults to 1
- _endRow_ the last row to import
- _startCol_ the starting column name (A, B etc) or column number (1, 2 etc.)
- _endCol_ the end column name (K, L etc) or column number (11, 12 etc.)
- _firstRowAsColNames_ whether the first row should be used for the names of each column, if false the column names will be v1, v2 etc. Defaults to true