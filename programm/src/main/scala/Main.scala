
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
    //Print the Head of the DataFrame == The first Line of the DataFrame
    println(df.head.mkString(", "))
    //Print the first 5 rows of the DataFrame with data going from the second line of the DataFrame
    df.slice(1, 6).foreach(row => println(row.mkString(", ")))
    //Sort the DataFrame by the first column (Date)
    //seperating the header from the data, so that the header is not sorted
    val header = df.head
    val data =  df.tail
    //Sort the DataFrame by the first column (Date) and putting them back together
    val sortedDF = header :: data.sortBy(row => row.head)
    println("Sorted DataFrame:")
    sortedDF.slice(0, 6).foreach(row => println(row.mkString(", ")))
    //Removing the column EM with the index 9
    val dfWithoutColumn = df.map(row => row.patch(9, Nil, 9))
    println(dfWithoutColumn.head.mkString(", "))
    // Get column name from header
    val columnName = df.head(6)
    // Find max value (excluding header)
    val maxValue = data.tail.map(row => row(6).toDouble).max
    println(s"Maximum value for $columnName: $maxValue")
    // Add new column by mapping each row
    val dfMean = df.map { row =>
      if (row == df.head) {
        // Add header for new column
        row :+ "Mean_DAX_FTSE"
      } else {
        // Add calculated value for new column
        val dax = row(3).toDouble
        val ftse = row(4).toDouble
        val mean = (dax + ftse) / 2
        row :+ mean.toString
      }
    }
    println("Header:")
    println(dfMean.head.mkString(", "))
    println("\nFirst 5 rows with mean:")
    dfMean.slice(1, 6).foreach(row => println(row.mkString(", ")))
  } catch {
    case e: Exception => println(s"Error reading file: ${e.getMessage}")
  }
}
