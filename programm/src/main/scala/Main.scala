
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import scala.collection.mutable.ListBuffer
import java.text.SimpleDateFormat
import org.apache.poi.ss.usermodel.DateUtil
object Main extends App {
  def readXlsxFile(filePath: String): List[List[String]] = {
    val file = new FileInputStream(filePath)
    val workbook = new XSSFWorkbook(file)
    val sheet = workbook.getSheetAt(0)
    val dateFormat = new SimpleDateFormat("yyyy-MM-dd")
  
    val rows = ListBuffer[List[String]]()
  
    val rowIterator = sheet.iterator()
    while (rowIterator.hasNext) {
      val currentRow = rowIterator.next()
      val cellIterator = currentRow.iterator()
    
      val cellValues = ListBuffer[String]()
      while (cellIterator.hasNext) {
        val cell = cellIterator.next()
        val cellValue = cell.getCellType match {
          case org.apache.poi.ss.usermodel.CellType.STRING => cell.getStringCellValue
          case org.apache.poi.ss.usermodel.CellType.NUMERIC if DateUtil.isCellDateFormatted(cell) => 
            dateFormat.format(cell.getDateCellValue)
          case org.apache.poi.ss.usermodel.CellType.NUMERIC => cell.getNumericCellValue.toString
          case org.apache.poi.ss.usermodel.CellType.BOOLEAN => cell.getBooleanCellValue.toString
          case _ => ""
        }
        cellValues += cellValue
      }
      rows += cellValues.toList
    }
  
    workbook.close()
    file.close()
  
    rows.toList
  }

// Example usage
  try {
    val df = readXlsxFile("src/main/resources/ISTANBUL STOCK EXCHANGE DATA SET.xlsx")
    print(df.head.mkString(", "))
  } catch {
    case e: Exception => println(s"Error reading file: ${e.getMessage}")
  }
}
