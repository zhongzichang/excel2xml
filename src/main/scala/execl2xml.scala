import java.io.File
import java.io.FileInputStream
import java.io.PrintWriter
import org.apache.poi.xssf.usermodel.{XSSFCell, XSSFRow, XSSFSheet, XSSFWorkbook}
import org.apache.poi.ss.usermodel.{Sheet,Row,Cell}
import scala.util.matching.Regex
import scala.collection.JavaConverters._
import scala.collection.mutable
import scala.xml

object Excel2Xml {

  

  def main(args: Array[String]) = {
    if( args.length == 2)
      excel2Xml(new File(args(0)),new File(args(1)))
    else
      println("usage: excel2xml <src> <tgt>")
  }

  def excel2Xml(src: File, dst:File): Unit = {
    if( !src.exists ){
      println("source file or directory not found.")
      return
    }
    if( !dst.exists ){
      println("target directory not found.")
      return
    }
    if( src exists ){
      val r = """[^~\.].*\.xlsx$""".r
      def h(f: File) = {transform(f, dst)}
      def recursiveHandleFiles(f: File): Unit = {
        if( f.isDirectory ) {
          f.listFiles.foreach(recursiveHandleFiles)
        } else if( r.findFirstIn(f.getName).isDefined ) {
          h(f)
        }
      }
      recursiveHandleFiles(src)
    }
  }

  def transform(f: File, dstDir: File): Unit = {

    def txSheet(s: Sheet) = {
      var curRow: Int = 0
      var ths: Array[String] = new Array(0);
      val sb: StringBuilder = StringBuilder.newBuilder
      def txRow(r: Row) = {
        var curCol: Int = 0
        def txCell(c: Cell) = {
          if( curRow > 1 ){
            c.getCellType match {
              case Cell.CELL_TYPE_BLANK => {
                sb ++= "<"+ths(curCol)+">"+"</"+ths(curCol)+">"
              }
              case Cell.CELL_TYPE_BOOLEAN => {
                val el = <e>{c.getBooleanCellValue}</e>
                sb ++= el.copy(label = ths(curCol)).toString
              }
              case Cell.CELL_TYPE_ERROR => {
                val el = <e>{c.getErrorCellValue}</e>
                sb ++= el.copy(label = ths(curCol)).toString
              }
              case Cell.CELL_TYPE_FORMULA => {
                sb ++= "<"+ths(curCol)+">"+"</"+ths(curCol)+">"
              }
              case Cell.CELL_TYPE_NUMERIC => {
                val el = <e>{c.getNumericCellValue}</e>
                sb ++= el.copy(label = ths(curCol)).toString
              }
              case Cell.CELL_TYPE_STRING => {
                val el = <e>{c.getStringCellValue}</e>
                sb ++= el.copy(label = ths(curCol)).toString
              }
            }
          } else if ( curRow == 0 ){
            ths(curCol) = c.getStringCellValue
          }
          curCol = curCol + 1
        }// txCell

        if( curRow == 0 ){
          ths = new Array(r.getLastCellNum)
          r.iterator.asScala.foreach(txCell)
        } else if( curRow > 1) {
          sb ++= "<value>"
          r.iterator.asScala.foreach(txCell)
          sb ++= "</value>"
        }
        curRow = curRow + 1
      }
      sb ++= "<root>"
      s.iterator.asScala.foreach(txRow)
      sb ++= "</root>"

      new PrintWriter(new File(dstDir, s.getSheetName + ".xml")){write(sb.toString); close}
    }


    val book = new XSSFWorkbook(new FileInputStream(f))
    val iterator = book.iterator.asScala
    iterator.foreach(txSheet)
  }


}
