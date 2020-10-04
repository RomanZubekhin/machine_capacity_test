import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class Test {
    private final static File fileRead = new File("C:\\Users\\user\\Desktop\\план.xlsx");
    private final static File fileWrite = new File("C:\\Users\\user\\Desktop\\нрм.xlsx");
    private int firstCellRead = 2;
    public static void main(String[] args) {
        int q = readFromExel(fileRead, 2, 3, 6);
        writeIntoExel(fileWrite, q);
    }

    public static int readFromExel(File file, int numArticle, int numName, int numQuantity) {
        int q = 0;
        XSSFWorkbook myExelBook = null;
        try {
            myExelBook = new XSSFWorkbook(new FileInputStream(file));
            myExelBook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        XSSFSheet myExelSheet = myExelBook.getSheet("Лист1");
        XSSFRow row = myExelSheet.getRow(1);

        if (row.getCell(numArticle).getCellType() == CellType.STRING) {
            System.out.println(row.getCell(numArticle).getStringCellValue());
        }
        if (row.getCell(numName).getCellType() == CellType.STRING){
            System.out.println(row.getCell(numName).getStringCellValue());
        }
        if (row.getCell(numQuantity).getCellType() == CellType.NUMERIC) {
            q = (int) row.getCell(numQuantity).getNumericCellValue();
        }
        return q;
    }
    public static void writeIntoExel(File fileWrite, int quantity){
        try {
            FileInputStream fileInputStream = new FileInputStream(fileWrite);
            Workbook wb = new XSSFWorkbook(fileInputStream);
            Sheet sheet = wb.getSheetAt(0);
            Cell cellQuantity = sheet.getRow(1).getCell(3);
            cellQuantity.setCellValue(quantity);

            //Re-evaluate formulas with POI's FormulaEvaluator
            wb.getCreationHelper().createFormulaEvaluator().evaluateAll();

            fileInputStream.close();
            FileOutputStream fileOutputStream = new FileOutputStream(fileWrite);
            wb.write(fileOutputStream);
            fileOutputStream.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}