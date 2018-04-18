# poi-collection 中文文档

[poi-collection](https://github.com/scalax/poi-collection) 是 Apache POI
Excel 部分 API 的一个轻量级封装，仅提供一些基础的 Scala 风格的 POI Excel API 封装。

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

| POI Cell | String Reader | Double Reader | Boolean Reader | Date Reader | Immutable String Reader | Non Empty String Reader | Non Blank String Reader |
|-------|-------|-------|-------|-------|-------|-------|-------|
| null | ""(empty string) | CellNotExistsException | CellNotExistsException | CellNotExistsException | ""(empty string) | CellNotExistsException | CellNotExistsException |
| Blank Cell | ""(empty string) | CellNotExistsException | CellNotExistsException | CellNotExistsException | ""(empty string) | CellNotExistsException | CellNotExistsException |
| StringCell("") | ""(empty string) | ExpectNumericCellException | ExpectBooleanCellException | ExpectDateException | ""(empty string) | CellNotExistsException | CellNotExistsException |
| StringCell("&nbsp;&nbsp;&nbsp;&nbsp;") | "&nbsp;&nbsp;&nbsp;&nbsp;" | ExpectNumericCellException | ExpectBooleanCellException | ExpectDateException | "&nbsp;&nbsp;&nbsp;&nbsp;" | "&nbsp;&nbsp;&nbsp;&nbsp;" | CellNotExistsException |
| StringCell("-123&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;") | "-123&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" | ExpectNumericCellException | ExpectBooleanCellException | ExpectDateException | "-123&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" | "-123&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" | "-123" |
| NumbericCell(123.321) | "123.321"(CellType->STRING) | 123.321 | ExpectBooleanCellException | Date(-2198535808600L) | ExpectStringCellException | "123.321"(CellType->STRING) | "123.321"(CellType->STRING) |
| NumbericCell(-123.321) | "-123.321"(CellType->STRING) | -123.321 | ExpectBooleanCellException | ExpectDateException | ExpectStringCellException | "-123.321"(CellType->STRING) | "-123.321"(CellType->STRING) |
| BooleanCell(true) | ExpectStringCellException | ExpectNumericCellException | true | ExpectDateException | ExpectStringCellException | ExpectStringCellException | ExpectStringCellException |
| BooleanCell(false) | ExpectStringCellException | ExpectNumericCellException | false | ExpectDateException | ExpectStringCellException | ExpectStringCellException | ExpectStringCellException |

这个表有几个要点：
* 对于 null 和 Blank Cell，使用 String Reader 将解析为空字符串，使用 Double Reader、Boolean Reader、Date Reader
将解析为 CellNotExistsException，这可能与 Apache POI 的默认行为有小许差异。
* 只有在 Numberic Cell 在使用 String Reader 解析的时候需要改变 CellType，使用 Immutable String Reader 解析可以避免这个问题，但将无法获取到
Numberic Cell 的内容。如果有其他 Formula Cell 依赖这个 Cell 来计算结果，使用 Immutable String Reader 可以避免 Cell 类型的改变。
* Non Empty String Reader 是在 String Reader 的基础上把所有的空字符串视为 CellNotExistsException。
* Non Blank String Reader 会把由 String Reader 解析到的字符串进行 trim 操作，然后依据 Non Empty String Reader 的行为继续解析，也就是说会把只有空格的字符串也解析为 CellNotExistsException。
* 对于 Formula Cell，对应的 Reader 将会先对 Cell 进行计算，然后作为普通 Cell 进行解析，计算过程会改变 Cell 的状态。
* 根据上面所描述的情况，在读取过程中，Cell 的状态（包括 CellType 和 Value）将会发生改变，应避免再使用此 Workbook 进行其他操作。

上述 Reader 都已扩展了 Option[T] 类型的 Reader。如已经引入了 String Reader，则可使用 tryValue[Option[String]] 进行解析。
Option 类的 Reader 将只把 CellNotExistsException 转化为 None，把正常返回值转化为 Option(value)，不会转化其他异常。

如需转化其他异常或者需要提供其他类型的 CellReader，只需
```scala
import cats.implicits._
```
poi-collection 已经实现了一个 MonadError[CellReader]，可自行对 CellReader 进行扩展。

### 写入

poi-collection 的写入依然使用了 Type Class 风格的封装。
这个写入封装在尽量保持 Scala 代码风格的同时减少了 CellStyle 的产生，
亦可避免遇到 HSSFWorkbook CellStyle 数量不能超过 4000 的问题。
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
注意：
* 所有继承自 StyleTransform 的 class 和 object 都必须为 case class 或 case object，这样可以更好地分辨重复的 CellStyle
处理链条。
* 不要使用参数中的 Workbook 创建 CellStyle，只需改参数中的 CellStyle 即可，但 Workbook 可用于创建 DataFormat 等对象。
* poi-collection 已经实现了一个 Contravariant[CellWriter]，导入 cats 的相关隐式转换后可以使用 contramap
方法扩展更多类型的 CellWriter。

然后使用以下代码产生副作用写入至 Workbook 即可：
```scala
val gen = StyleGen.getInstance
CPoiUtils.multiplySet(gen, cells): StyleGen
```
CPoiUtils.multiplySet 的返回值是一个新的 StyleGen，拥有设值过程中产生的 CellStyle 缓存，如果在一组设值操作中有多段设值代码，
为了充分使用上一个设值操作的 CellStyle 缓存，可以继续使用
CPoiUtils.multiplySet 的返回值作为下一个 CPoiUtils.multiplySet 的 gen 参数。

在性能敏感的场合，可以使用以下方法进行设值操作，MutableStyleGen 将会使用 mutable.Map 来记录 CellStyle 处理链的缓存。
```scala
val gen = MutableStyleGen.getInstance
CPoiUtils.multiplySet(gen, cells): Unit
```
第一句定义的 gen 可以重复使用在同一个 Workbook 的设值操作中以充分利用 CellStyle 缓存。

注意：MutableStyleGen 不是线程安全的，但并不影响最终效果。MutableStyleGen
只是为了缩减大量重复的 CellStyle，并发有可能会造成 CellStyle 数量的少量增加，但并不会造成 CellStyle 数量的暴涨。
