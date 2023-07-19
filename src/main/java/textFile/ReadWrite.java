package textFile;

import org.apache.commons.io.FileUtils;

import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;



public class ReadWrite {
    public static void main(String[] args) {

            String filePath = "src/main/resources/sample.txt";

            // Read and display the content from the text file
            readAndDisplayContent(filePath);

            // Add new text in a new line and write back to the file
            String newLine = "\nThe total amount would be $110.";
            appendTextToFile(filePath, newLine);
        }

        private static void readAndDisplayContent(String filePath) {
            try {
                String content = FileUtils.readFileToString(new File(filePath), StandardCharsets.UTF_8);
                System.out.println("File Content:");
                System.out.println(content);
            } catch (IOException e) {
                System.err.println("Error reading the file: " + e.getMessage());
            }
        }

        private static void appendTextToFile(String filePath, String text) {
            try {
                FileUtils.writeStringToFile(new File(filePath), text, StandardCharsets.UTF_8, true);
                System.out.println("Text has been added to the file successfully.");
            } catch (IOException e) {
                System.err.println("Error writing to the file: " + e.getMessage());
            }
        }
    }
