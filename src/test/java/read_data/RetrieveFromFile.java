package read_data;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;

import java.io.FileNotFoundException;
import java.io.IOException;

public class RetrieveFromFile {
    @Test
    public void readFileTest() throws IOException {
        File excelFile = new File("src/test/resources/TestSetup.xlsx");
        FileInputStream fileInputStream = new FileInputStream(excelFile);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet page1 = workbook.getSheet("Sheet1");
               XSSFRow row1= page1.getRow(0);
                     XSSFCell cell1=  row1.getCell(0);

        System.out.println(cell1);
    }
    @Test
    public void getRowValuesTest() throws IOException {
        File excelFile = new File("src/test/resources/TestSetup.xlsx");
        FileInputStream fileInputStream = new FileInputStream(excelFile);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet1 = workbook.getSheetAt(0);
        XSSFRow row1 = sheet1.getRow(0);

        for (int i = row1.getFirstCellNum(); i < row1.getLastCellNum() ; i++) {
            XSSFCell cell = row1.getCell(i);
            System.out.print(cell + " | ");

        }
    }
    @Test
    public void getAllDataTest() throws IOException {
        //Get all data from excell document
        File File = new File("src/test/resources/TestSetup.xlsx");
        FileInputStream fileInputStream = new FileInputStream(File);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);


        XSSFSheet sheet = workbook.getSheetAt(0);


        for (int i = sheet.getFirstRowNum(); i < sheet.getLastRowNum() ; i++) {
            XSSFRow temprow = sheet.getRow(i);
            for (int j = temprow.getFirstCellNum(); j<temprow.getLastCellNum() ; j++) {

                System.out.print(temprow.getCell(j) + " | ");

            }
            System.out.println();

        }

    }
}
