# poi-collection 中文文档

[poi-collection](https://github.com/scalax/poi-collection) 是 Apache POI
Excel 部分一个轻量级的封装。仅提供一些基础的 Scala 风格的 POI Excel API 封装。

## 关注点

poi-collection 主要为了解决以下两个问题：
* 类型友好的 Type Class 风格读写器。
* Scala 友好的 CellStyle 数量控制方案。

## 使用

1.读取

```scala
import net.scalax.cpoi._
import readers._

val file = new File("filePath")
val outputStream = new FileOutputStream(file)
val workbook = new HSSFWorkbook(outputStream)
val sheet = workbook.getSheet("Sheet1")
val row = sheet.getRow(2)
val cell1 = row.getCell(4)
val result1: Either[CellReaderException, String] = CPoiUtils.wrapCell(cell1).tryValue[String] //Right("Test")
val cell2 = row.getCell(5) //Boolean Cell
val result2: Either[CellReaderException, String] = CPoiUtils.wrapCell(cell2).tryValue[String] //Left(ExpectStringCellException)
val cell3 = row.getCell(6) //null
val result3: Either[CellReaderException, String] = CPoiUtils.wrapCell(cell3).tryValue[Option[Double]] //Right(None)
```

CPoiUtils.wrapCell 的使用方法如下：
```scala
CPoiUtils.wrapCell(cell)
CPoiUtils.wrapCell(null: Cell)
CPoiUtils.wrapCell(Option(cell))
```

tryValue 的行为可能与 POI 的默认行为有些差异，以下为 tryValue 的具体行为列表：

| POI Cell | String reader | Double reader | Boolean reader | Date reader | Immutable string reader | Non empty string reader | Non blank string reader |
|-------|-------|-------|-------|-------|-------|-------|-------|
| null | ""(empty string) | CellNotExistsException | CellNotExistsException | CellNotExistsException | ""(empty string) | CellNotExistsException | CellNotExistsException |
| Blank Cell | ""(empty string) | CellNotExistsException | CellNotExistsException | CellNotExistsException | ""(empty string) | CellNotExistsException | CellNotExistsException |
| StringCell(""(empty string)) | ""(empty string) | ExpectNumericCellException | ExpectBooleanCellException | ExpectDateException | ""(empty string) | CellNotExistsException | CellNotExistsException |
| StringCell("&nbsp;&nbsp;&nbsp;&nbsp;") | "&nbsp;&nbsp;&nbsp;&nbsp;" | ExpectNumericCellException | ExpectBooleanCellException | ExpectDateException | "&nbsp;&nbsp;&nbsp;&nbsp;" | "&nbsp;&nbsp;&nbsp;&nbsp;" | CellNotExistsException |
| StringCell("-123&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;") | "-123&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" | ExpectNumericCellException | ExpectBooleanCellException | ExpectDateException | "-123&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" | "-123&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" | "-123" |
| NumbericCell(123.321) | "123.321"(CellType to STRING) | 123.321 | ExpectBooleanCellException | Date(-2198535808600L) | ExpectStringCellException | "123.321"(CellType to STRING) | "123.321"(CellType to STRING) |
| NumbericCell(-123.321) | "-123.321"(CellType to STRING) | -123.321 | ExpectBooleanCellException | ExpectDateException | ExpectStringCellException | "-123.321"(CellType to STRING) | "-123.321"(CellType to STRING) |
| BooleanCell(true) | ExpectStringCellException | ExpectNumericCellException | true | ExpectDateException | ExpectStringCellException | ExpectStringCellException | ExpectStringCellException |
| BooleanCell(false) | ExpectStringCellException | ExpectNumericCellException | false | ExpectDateException | ExpectStringCellException | ExpectStringCellException | ExpectStringCellException |

上述 Reader 都已扩展了 Option[T] 类型的 Reader。如已经 import 了 String reader，则可使用 tryValue[Option[String]]。
Option 类的 Reader 将只把 CellNotExistsException 转化为 None，把正常返回值转化为 Option(value)，不会转化其他异常。

如需转化其他异常或者需要提供其他类型的 Reader，只需
```scala
import cats.implicits._
```
poi-collection 已经提供了 CellReader 类型的 MonadError ，可自行扩展该 Reader。