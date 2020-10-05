import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;
//import java.util.Map;

public class Test {
    private final static File fileRead = new File("C:\\Users\\user\\Desktop\\план.xlsx");
    //private final static File fileWrite = new File("C:\\Users\\user\\Desktop\\нрм.xlsx");
    private final static HashMap<String,Integer> hashMap = new HashMap<String, Integer>();

    public static void main(String[] args) throws IOException {
        readFromExel(fileRead, 2, 6);
        //print(hashMap);
        //writeIntoExel(fileWrite, 2);
    }

    public static void readFromExel(File file, int numArticle, int numQuantity) throws IOException {
        int count = 1;
        Workbook myExelBook = new XSSFWorkbook(new FileInputStream(file));
        Sheet myExelSheet = myExelBook.getSheet("Лист1");
        while (true) {
            try {
                Row row = myExelSheet.getRow(count);
                String article = null;
                int quantity = 0;
                Cell c = row.getCell(numArticle);
                if (!(c == null || c.getCellType() == CellType.BLANK)) {
                    if (row.getCell(numArticle).getCellType() == CellType.STRING) {
                        article = row.getCell(numArticle).getStringCellValue();
                    }
                    if (row.getCell(numQuantity).getCellType() == CellType.NUMERIC) {
                        quantity = (int) row.getCell(numQuantity).getNumericCellValue();
                    }
                    hashMap.put(article, quantity);
                    count++;
                } else break;
            }catch (NullPointerException exception){return;}
        }
    }

    public static void writeIntoExel(File fileWrite, HashMap<String, Integer> map){
        try {
            FileInputStream fileInputStream = new FileInputStream(fileWrite);
            Workbook wb = new XSSFWorkbook(fileInputStream);
            Sheet sheet = wb.getSheetAt(0);
            int startCellColumn = 2;
            int startCellRow = 1;
            Cell cell= sheet.getRow(1).getCell(3);
            cell.setCellValue(3);

            //Re-evaluate formulas with POI's FormulaEvaluator
            wb.getCreationHelper().createFormulaEvaluator().evaluateAll();
            fileInputStream.close();
            //write data
            FileOutputStream fileOutputStream = new FileOutputStream(fileWrite);
            wb.write(fileOutputStream);
            fileOutputStream.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

/*
    public static void print(HashMap<String,Integer> map){
        String article;
        int quantity;
        int count = 1;
        for (Map.Entry<String, Integer> pair : map.entrySet()) {
            article = pair.getKey();
            quantity = pair.getValue();
            System.out.println(count + ") Article = " + article + " " + "Quantity = " + quantity);
            count++;
        }
    }
*/
}