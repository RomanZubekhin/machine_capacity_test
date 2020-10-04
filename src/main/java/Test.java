import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

public class Test {
    private final static File fileRead = new File("C:\\Users\\user\\Desktop\\план.xlsx");
    private final static File fileWrite = new File("C:\\Users\\user\\Desktop\\нрм.xlsx");

    public static void main(String[] args) throws IOException {
        readFromExel(fileRead, 2, 6);
        print(hashMap);
        //writeIntoExel(fileWrite, 2);

    }
    private static HashMap<String,Integer> hashMap = new HashMap<String, Integer>();

    public static void readFromExel(File file, int numArticle, int numQuantity) throws IOException {
        System.out.println("Enter to method readFromExel");
        int count = 1;
        Workbook myExelBook = new XSSFWorkbook(new FileInputStream(file));
        Sheet myExelSheet = myExelBook.getSheet("Лист1");
        while (true) {
            Row row = myExelSheet.getRow(count);
            String article = null;
            int quantity = 0;
            if(!(row.getCell(numArticle) == null)) {
                if (row.getCell(numArticle).getCellType() == CellType.STRING) {
                    article = row.getCell(numArticle).getStringCellValue();
                }
                if (row.getCell(numQuantity).getCellType() == CellType.NUMERIC) {
                    quantity = (int) row.getCell(numQuantity).getNumericCellValue();
                }
                hashMap.put(article, quantity);
                count++;
            }

        }
        //myExelBook.close();
        //System.out.println("Exit to method readFromExel");
    }

    public static void print(HashMap<String,Integer> map){
        System.out.println("Enter to method print");

        String article;
        int quantity;
        for (Map.Entry<String, Integer> pair : map.entrySet()) {
            article = pair.getKey();
            quantity = pair.getValue();
            System.out.println("Article= " + article + "Quantity=" + quantity);
        }
        System.out.println("Exit to method print");
    }
   /* public static void writeIntoExel(File fileWrite, int quantity){
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
    }*/
}