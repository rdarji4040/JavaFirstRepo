package textFile;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Scanner;

public class StudentExcelFile {
    public static void main(String[] args) {
        try {
            // Read the .xlsx / .xls file
            File file = new File("C:\\Users\\rdarj\\Downloads\\students.xlsx");
            FileInputStream fileInputStream = new FileInputStream(file);

            // Create workbook instance and capture the data from the excel file
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            // Read the first sheet
            XSSFSheet xssfSheet = workbook.getSheetAt(0);

            // Display all the student information as an array of Student objects
            ArrayList<Student> studentList = new ArrayList<>();
            // ... (existing code)

// Read data from the Excel file and populate the studentList
            for (int rowIndex = 1; rowIndex <= xssfSheet.getLastRowNum(); rowIndex++) {
                XSSFRow row = xssfSheet.getRow(rowIndex);
                String studentId = row.getCell(0).getStringCellValue(); // Assuming student ID is a string
                String name = row.getCell(1).getStringCellValue(); // Assuming name is a string
                int score1 = (int) row.getCell(2).getNumericCellValue(); // Assuming scores are numeric
                int score2 = (int) row.getCell(3).getNumericCellValue();
                int score3 = (int) row.getCell(4).getNumericCellValue();

                Student student = new Student(studentId, name, score1, score2, score3);
                studentList.add(student);
            }

// ... (existing code)


            // Display the student information as an array of Student objects
            for (Student student : studentList) {
                System.out.println(student.getStudentId() + " | " + student.getName() + " | " +
                        student.getScore1() + " | " + student.getScore2() + " | " + student.getScore3() + " | " +
                        student.getFinalScore());
            }

            // Prompt the user to enter data for 2 new students
            Scanner scanner = new Scanner(System.in);
            for (int i = 0; i < 2; i++) {
                System.out.println("Enter Student ID:");
                String studentId = scanner.next();

                System.out.println("Enter Student Name:");
                String name = scanner.next();

                System.out.println("Enter Score 1:");
                int score1 = scanner.nextInt();

                System.out.println("Enter Score 2:");
                int score2 = scanner.nextInt();

                System.out.println("Enter Score 3:");
                int score3 = scanner.nextInt();

                Student newStudent = new Student(studentId, name, score1, score2, score3);
                studentList.add(newStudent);
            }

            // Write the new student data to the existing Excel file
            int rowIndex = xssfSheet.getLastRowNum() + 1;
            for (Student student : studentList) {
                Row row = xssfSheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(student.getStudentId());
                row.createCell(1).setCellValue(student.getName());
                row.createCell(2).setCellValue(student.getScore1());
                row.createCell(3).setCellValue(student.getScore2());
                row.createCell(4).setCellValue(student.getScore3());
                row.createCell(5).setCellValue(student.getFinalScore());
            }

            // Write back to Excel
            FileOutputStream fileOutputStream = new FileOutputStream(file);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            System.out.println("Successfully Saved to the excel file!!");

        } catch (Exception e) {
            System.err.println("Error is: " + e.getMessage());
        }
    }
}
