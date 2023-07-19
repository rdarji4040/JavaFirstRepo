package textFile;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

public class ReadExcelFile {
    public static void main(String[] args) {
        try {
            // read the .xls file
            File file = new File("C:\\Users\\rdarj\\OneDrive\\Desktop\\Projects\\NewJavaProject\\src\\main\\resources\\sample.text");

            FileInputStream fileInputStream = new FileInputStream(file) ;
            // create workbook instance and capture the data from excel file
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            int noOfSheet= workbook.getNumberOfSheets();
            System.out.println(" number of sheet present in xls file is: " + noOfSheet);
            // read the first sheet
            XSSFSheet xssfSheet = workbook.getSheetAt(0);
            // read the rows one by one
            Iterator<Row> rowIterator = xssfSheet.iterator();
            while (rowIterator.hasNext()){
                Row row = rowIterator.next();
                // fetch all the column
                Iterator <Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()){
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()){
                        case STRING:
                            System.out.println(cell.getStringCellValue() + "/t");
                            break;
                        case NUMERIC:
                            System.out.println(cell.getNumericCellValue() + "/t");
                            break;
                        case BOOLEAN:
                            System.out.println(cell.getBooleanCellValue() + "/t");

                    }
                }
            }
        }
        catch (Exception e){
            System.out.println("error is " + e.getMessage() );
        }
    }

}
