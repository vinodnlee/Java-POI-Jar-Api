package JavaExcelPoi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

/**
 * Created with IntelliJ IDEA.
 * User: Ragavi
 * Date: 4/20/14
 * Time: 5:52 AM
 * To change this template use File | Settings | File Templates.
 */
public class ReadExcel {

    public static void main (String a[])
    {
        try{
        FileInputStream file = new FileInputStream(new File("L:\\1.Knowledge_Vinod\\2.SOFTWARE_VIDEO\\1.Learn _java\\JAVA Video traning\\javatopcis\\javaExcelUsingPoi\\JavaPoiExamples\\Sample.xls"));
        Workbook workbook = new HSSFWorkbook(file);
        Sheet sheet =  workbook.getSheetAt(0);
        for(Row row : sheet){
            for(Cell cell : row){
              cell.setCellType(Cell.CELL_TYPE_STRING);
              System.out.print(cell.getStringCellValue());
            }
            System.out.println();
        }
        }
        catch (Exception e){
            e.printStackTrace();
        }

    }
}
