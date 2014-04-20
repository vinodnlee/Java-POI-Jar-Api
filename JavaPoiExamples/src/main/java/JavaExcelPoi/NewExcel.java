package JavaExcelPoi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;

/**
 * Created with IntelliJ IDEA.
 * User: Ragavi
 * Date: 4/20/14
 * Time: 1:39 AM
 * To change this template use File | Settings | File Templates.
 */
public class NewExcel {
    public static void main(String []args){
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("testSheet");
        Cell cell =  sheet.createRow(0).createCell(3);
        cell.setCellValue("Hi there");
        cell.getRow().setHeightInPoints(30);
        sheet.setColumnWidth(3,7000);
        System.out.println(cell.getRichStringCellValue().toString());

        //Cell formulas
        Cell cell1 = sheet.createRow(1).createCell(0);
        Cell cell2 = sheet.createRow(1).createCell(1);
        Cell cell3 = sheet.createRow(1).createCell(2);
        Cell cell4 = sheet.createRow(1).createCell(3);
        Cell cell5 = sheet.createRow(1).createCell(4);
        cell1.setCellValue(10);
        cell2.setCellValue(10);
        cell3.setCellValue(10);
        cell4.setCellValue(10);
        cell5.setCellFormula("SUM(A2:D2)");

        //Cell Style
        CellStyle style = workbook.createCellStyle();
        style.setFillBackgroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        Font font =  workbook.createFont();
        font.setColor(IndexedColors.YELLOW.getIndex());
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        font.setItalic(true);
        font.setFontHeightInPoints((short)16);
        font.setUnderline(Font.U_DOUBLE);
        font.setFontName("Helvetia");
        style.setFont(font);
        cell.setCellStyle(style);
        cell.setCellValue("Vinod");

        //iterate xl Document

        try{
            FileOutputStream output = new FileOutputStream("Test.xls");
            workbook.write(output);
            output.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}
