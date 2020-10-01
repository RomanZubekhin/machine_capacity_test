import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class Test {
    private final static File fileRead = new File("C:\\Users\\user\\Desktop\\план.xlsx");
    private final static File fileWrite = new File("C:\\Users\\user\\Desktop\\нрм.xlsx");
    public static void main(String[] args) {
        readFromExel(fileRead, 2, 3, 6);
        writeIntoExel(fileWrite);
    }

    public static void readFromExel(File file, int numArticle, int numName, int numQuantity) {
        try {
            XSSFWorkbook myExelBook = new XSSFWorkbook(new FileInputStream(file));
            XSSFSheet myExelSheet = myExelBook.getSheet("Лист1");
            XSSFRow row = myExelSheet.getRow(1);

            if (row.getCell(numArticle).getCellType() == CellType.STRING) {
                System.out.println(row.getCell(numArticle).getStringCellValue());
            }
            if (row.getCell(numName).getCellType() == CellType.STRING){
                System.out.println(row.getCell(numName).getStringCellValue());
            }
            if (row.getCell(numQuantity).getCellType() == CellType.NUMERIC) {
                System.out.println(row.getCell(numQuantity).getNumericCellValue());
            }
            myExelBook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    public static void writeIntoExel(File fileWrite){
        try {
            FileInputStream fileInputStream = new FileInputStream(fileWrite);
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
            XSSFCell cell = sheet.getRow(1).getCell(3);
            cell.setCellValue("add value");
            fileInputStream.close();

            FileOutputStream fileOutputStream = new FileOutputStream(fileWrite);
            xssfWorkbook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}