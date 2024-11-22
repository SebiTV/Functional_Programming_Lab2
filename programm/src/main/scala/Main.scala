
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import scala.collection.mutable.ListBuffer
import java.text.SimpleDateFormat
import org.apache.poi.ss.usermodel.DateUtil
import java.io.FileOutputStream

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

  def writeToXlsx(data: List[List[String]], path: String): Unit = {
    val workbook = new XSSFWorkbook()
    val sheet = workbook.createSheet("Sheet1")
    
    data.zipWithIndex.foreach { case (row, rowIndex) =>
      val excelRow = sheet.createRow(rowIndex)
      row.zipWithIndex.foreach { case (cellValue, cellIndex) =>
        val cell = excelRow.createCell(cellIndex)
        cell.setCellValue(cellValue)
      }
    }
    
    val output_file = new FileOutputStream(path)
    workbook.write(output_file)
    workbook.close()
    output_file.close()
  }

  try {
    val df = readXlsxFile("src/main/resources/ISTANBUL STOCK EXCHANGE DATA SET.xlsx")
    //Print the Head of the DataFrame == The first Line of the DataFrame
    println(df.head.mkString(", "))
    //Print the first 5 rows of the DataFrame with data going from the second line of the DataFrame
    df.slice(1, 6).foreach(row => println(row.mkString(", ")))

    //Sort the DataFrame by the first column (Date)
    //seperating the head from the data, so that the head is not sorted
    val head = df.head
    val data =  df.tail

    //Sort the DataFrame by the first column (Date) and putting them back together
    val sortedDF = head :: data.sortBy(row => row.head).reverse
    println("Sorted DataFrame:")
    sortedDF.slice(0, 6).foreach(row => println(row.mkString(", ")))

    //Removing the column EM with the index 9
    val dfWithoutColumn = sortedDF.map(row => row.patch(9, Nil, 9))
    println(dfWithoutColumn.head.mkString(", "))

    //Get the column name from head
    val columnName = dfWithoutColumn.head(6)
    //Error handling for numeric conversions
    def safeToDouble(s: String): Double = {
      try {
        s.toDouble
      } catch {
        case _: NumberFormatException => 0.0
      }
    }
    val head2 = dfWithoutColumn.head
    val data2 = dfWithoutColumn.tail
    val sortedDF2 = head2 :: data2.sortBy(row => row(6).toDouble).reverse
    
    //Add new column by mapping each row
    val dfMean = sortedDF2.map { row =>
      if (row == sortedDF2.head) {
        //Add head for new column with the name of the column
        row :+ "Mean_DAX_FTSE"
      } else {
        //Add calculated value for new column
        val dax = safeToDouble(row(3))
        val ftse = safeToDouble(row(4))
        val mean = (dax + ftse) / 2
        row :+ mean.toString
      }
    }
    //Print the head and the first 5 rows with the new columnc
    println("head:")
    println(dfMean.head.mkString(", "))
    println("First 5 rows with mean:")
    dfMean.slice(1, 6).foreach(row => println(row.mkString(", ")))
    def find_highest_value(columnNumber: Int): (Double) = {
      val maxValue = data.tail.map(row => row(columnNumber).toDouble).max
      maxValue
    }
    println(s"Highest value for column 6: ${find_highest_value(6)}")
    writeToXlsx(dfMean, "src/main/resources/output.xlsx")
  } catch {
    case e: Exception => println(s"Error: ${e.getMessage}")
  }
}
