package textFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Scanner;
public class StudentExcelFile {
    int studentId;
    String studentName;
    double sub1Score;
    double sub2Score;
    double sub3Score;
    double finalScore;

    public static void main(String[] args) {
        try {
            // Creating file object of existing Excel file
            File studentExcelFile = new File("C:\\Users\\rdarj\\Downloads\\students.xlsx");
            FileInputStream fileInputStream = new FileInputStream(studentExcelFile);  //Creating input stream
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);  //Creating workbook from input stream

            int noOfSheet = workbook.getNumberOfSheets();       //Reading first sheet of Excel file
            System.out.println("Number of sheets present in xlsx file is " + noOfSheet);
            XSSFSheet xssfSheet = workbook.getSheetAt(0); //Read the first sheet
            int rowCount = xssfSheet.getLastRowNum();        //Getting the count of existing records
            System.out.println(xssfSheet.getLastRowNum());  // gives us the last row number

            Scanner scanner = new Scanner(System.in); //ask usersInput and students detail
            System.out.print("Enter the number of new students to add: ");
            int numStudents = scanner.nextInt();
            scanner.nextLine();  // Consume the remaining newline character after reading the integer

            for (int i = 0; i < numStudents; i++) {
                System.out.println("Enter details for Student " + (i + 1) + ":");
                System.out.print("studentId: ");
                int studentId = scanner.nextInt();
                scanner.nextLine(); // Consume the remaining newline character after reading the integer
                System.out.print("studentName: ");
                String studentName = scanner.nextLine();
                System.out.print("sub1Score: ");
                double sub1Score = scanner.nextDouble();
                System.out.print("sub2Score: ");
                double sub2Score = scanner.nextDouble();
                System.out.print("sub3Score: ");
                double sub3Score = scanner.nextDouble();
                System.out.print("finalScore: ");
                double finalScore = scanner.nextDouble();

                // Create new row from the next row count and Set values for each cell in the new row
                Row row = xssfSheet.createRow(++rowCount);
                row.createCell(0).setCellValue(studentId);
                row.createCell(1).setCellValue(studentName);
                row.createCell(2).setCellValue(sub1Score);
                row.createCell(3).setCellValue(sub2Score);
                row.createCell(4).setCellValue(sub3Score);
                row.createCell(5).setCellValue(finalScore);
            }

            // 3. Update sub2_score from 89 to 93 where student id is 1002
            int targetStudentId1 = 1002;
            for (Row row : xssfSheet) {
                Cell idCell = row.getCell(0);
                if (idCell != null && idCell.getCellType() == CellType.NUMERIC) {
                    int studentId = (int) idCell.getNumericCellValue();
                    if (studentId == targetStudentId1) {
                        Cell sub2ScoreCell = row.getCell(3);
                        if (sub2ScoreCell != null && sub2ScoreCell.getCellType() == CellType.NUMERIC) {
                            sub2ScoreCell.setCellValue(93);
                        }
                    }
                }
            }

// 4. Update student name to 'Tiger' where student id is 1005
            int targetStudentId2 = 1005;
            for (Row row : xssfSheet) {
                Cell idCell = row.getCell(0);
                if (idCell != null && idCell.getCellType() == CellType.NUMERIC) {
                    int studentId = (int) idCell.getNumericCellValue();
                    if (studentId == targetStudentId2) {
                        Cell studentNameCell = row.getCell(1);
                        if (studentNameCell != null && studentNameCell.getCellType() == CellType.STRING) {
                            studentNameCell.setCellValue("Tiger");
                        }
                    }
                }
            }

            // 5. Update final_score to 81 where name is 'Jack'
            String targetStudentName = "Jack";
            for (Row row : xssfSheet) {
                Cell nameCell = row.getCell(1);
                if (nameCell != null && nameCell.getCellType() == CellType.STRING) {
                    String studentName = nameCell.getStringCellValue();
                    if (studentName.equals(targetStudentName)) {
                        Cell finalScoreCell = row.getCell(5);
                        if (finalScoreCell != null && finalScoreCell.getCellType() == CellType.NUMERIC) {
                            finalScoreCell.setCellValue(81);
                        }
                    }
                }
            }


            fileInputStream.close();    //Close input stream

            //Creating output stream and writing the updated workbook
            FileOutputStream fileOutputStream = new FileOutputStream("C:\\Users\\rdarj\\Downloads\\students.xlsx");
            workbook.write(fileOutputStream);

            workbook.close();   //Close the workbook
            fileOutputStream.close(); //Close the output stream
            System.out.println("Excel file has been updated successfully.");

        } catch (Exception e) {
            System.err.println("Exception while updating an existing excel file: " + e.getMessage());
        }
    }

    public void Students(int studentId, String studentName, double sub1Score, double sub2Score, double sub3Score, double finalScore) {
        this.studentId = studentId;
        this.studentName = studentName;
        this.sub1Score = sub1Score;

        this.sub2Score = sub2Score;
        this.sub3Score = sub3Score;
        this.finalScore = finalScore;
    }
}
