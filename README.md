# poi-collection 中文文档

[poi-collection](https://github.com/scalax/poi-collection) 是 Apache POI
Excel 部分一个轻量级的封装。仅提供一些基础的 Scala 风格的 POI Excel API 封装。

## 关注点

poi-collection 主要为了解决以下两个问题：
* 类型友好的 Type Class 风格读写器。
* Scala 友好的 CellStyle 数量控制方案。

## 使用

### 读取

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

### 写入

poi-collection 的写入依然使用了 Type Class 风格的封装。
这个写入封装可以在尽量保持 Scala 代码风格的同时减少 cellStyle 的产生。
以避免遇到 HSSFWorkbook 4000 个 cellStyle 数量上限的问题。
如下则可建立一个 CellData：
```scala
case object TextStyle extends StyleTransform {
  override def operation(workbook: Workbook,
                       cellStyle: CellStyle): CellStyle = {
    val format = workbook.createDataFormat.getFormat("@")
    cellStyle.setDataFormat(format)
    cellStyle
  }
}

case object DoubleStyle extends StyleTransform {
  override def operation(workbook: Workbook,
                       cellStyle: CellStyle): CellStyle = {
    val format = workbook.createDataFormat.getFormat("0.00")
    cellStyle.setDataFormat(format)
    cellStyle
  }
}

case class Locked(lock: Boolean) extends StyleTransform {
  override def operation(workbook: Workbook,
                       cellStyle: CellStyle): CellStyle = {
    cellStyle.setLocked(lock)
    cellStyle
  }
}
  
import writers._
val cells = List(
  cell1 -> CellData(testUTF8Str).addTransform(TextStyle, Locked(false)),
  cell2 -> CellData(testDouble).addTransform(DoubleStyle, Locked(true))
)
```
注意：1、所有继承自 StyleTransform 的 class 和 object 都必须为 case class 或 case object 以便更好地分辨重复的 cellStyle
处理链条。

2、不要使用 workbook 创建 cellStyle，只需改变原 cellStyle 即可。

然后使用以下代码产生副作用作用于 Workbook 即可：
```scala
val gen = StyleGen.getInstance
CPoiUtils.multiplySet(gen, cells): StyleGen
```
CPoiUtils.multiplySet 的返回值是一个新的 styleGen，拥有设值过程中产生的 cellStyle 缓存，如果在一组设值操作中有多段设值代码，
可以继续使用 CPoiUtils.multiplySet 的返回值作为下一个 CPoiUtils.multiplySet 的 gen 参数以继续使用之前的 cellStyle 缓存。

如果对性能比较敏感，可以使用以下方法产生副作用，下面的方法将会使用 mutable.Map 来记录 cellStyle 处理链的缓存。
```scala
val gen = MutableStyleGen.getInstance
CPoiUtils.multiplySet(gen, cells): Unit
```
第一句定义的 gen 可以重复使用在同一个 workbook 的设值操作中以充分利用 cellStyle 缓存。

注意：MutableStyleGen 不是线程安全的，但并不影响最终效果。MutableStyleGen 只是为了缩减大量因为使用了 case class
声明方式而导致的重复 cellStyle，并发有可能会造成 cellStyle 数量的增加但并不会造成 cellStyle 数量的暴涨。
