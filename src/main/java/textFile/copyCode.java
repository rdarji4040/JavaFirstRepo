package textFile;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class copyCode {

        public static void main(String[] args) {
            try {

                //Read the .xlsx / .xls file
                File file = new File("C:\\Users\\rdarj\\Downloads\\students.xlsx");
                FileInputStream fileInputStream = new FileInputStream(file);

                //Create workbook instance and capture the data from excel file
                XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
                //Read the first sheet
                XSSFSheet xssfSheet = workbook.getSheetAt(0);
                System.out.println(xssfSheet.getLastRowNum());

                //create a new Row
                for (int j = 0; j < 2; j++) {
                    Row row = xssfSheet.createRow(xssfSheet.getLastRowNum() + 1);
                    for (int i = 0; i < 3; i++) {
                        Cell cell = row.createCell(i);
                        switch (i) {
                            case 0:
                                if (j == 0) {
                                    cell.setCellValue("F");
                                } else if (j == 1) {
                                    cell.setCellValue("G");
                                }
                                break;
                            case 1:
                                if (j == 0) {
                                    cell.setCellValue(15);
                                } else if (j == 1) {
                                    cell.setCellValue(14);
                                }
                                break;
                            case 2:
                                if (j == 0) {
                                    cell.setCellValue(5);
                                } else if (j == 1) {
                                    cell.setCellValue(7);
                                }
                                break;
                        }
                    }
                }

                //write back to excel
                FileOutputStream fileOutputStream = new FileOutputStream(file);
                workbook.write(fileOutputStream);
                fileOutputStream.close();
                System.out.println("Successfully Saved to the excel file!!");

            } catch (Exception e) {
                System.err.println("Error is :" +e.getMessage());
            }
        }
    }

