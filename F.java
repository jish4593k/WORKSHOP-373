import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class ExcelReaderWriter {

    public static void main(String[] args) {
        if (args.length > 0) {
            String excelFilename = args[0];
            try {
                System.out.println("Opening: " + excelFilename);
                Map<Integer, Map<String, Object>> spreadsheetData = readExcel(excelFilename);
                System.out.println("Data read from Excel:");
                System.out.println(spreadsheetData);

                // Let's assume you have some scraped data (dummy data for example)
                Map<Integer, String> scrapedData = new HashMap<>();
                scrapedData.put(1, "Scraped Info 1");
                scrapedData.put(2, "Scraped Info 2");

                writeScrapedData(spreadsheetData, scrapedData, "output.xlsx");
                System.out.println("Scraped data written to Excel.");

            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            System.out.println("Make sure you pass a file name on the command line --- syntax: java ExcelReaderWriter PMP_RFQ_FORM.xlsx");
        }
    }

    private static Map<Integer, Map<String, Object>> readExcel(String filename) throws IOException {
        Map<Integer, Map<String, Object>> partNumberData = new HashMap<>();
        try (FileInputStream fis = new FileInputStream(filename);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            for (int rowNumber = 1; rowNumber < sheet.getPhysicalNumberOfRows(); rowNumber++) {
                Row row = sheet.getRow(rowNumber);

                int lineNumber = (int) row.getCell(0).getNumericCellValue();
                // Extract other data fields as needed

                // Populate partNumberData map
                // ...

            }
        }
        return partNumberData;
    }

    private static void writeScrapedData(Map<Integer, Map<String, Object>> partNumberData,
                                         Map<Integer, String> scrapedData, String outputFilename) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Scraped Data");

            // Assuming you have header cells, add them here
            // ...

            int rowNum = 0;
            for (Map.Entry<Integer, String> entry : scrapedData.entrySet()) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(entry.getKey());
                row.createCell(1).setCellValue(entry.getValue());
            }

            // Write the workbook to an output file
            try (FileOutputStream fos = new FileOutputStream(outputFilename)) {
                workbook.write(fos);
            }
        }
    }
}
