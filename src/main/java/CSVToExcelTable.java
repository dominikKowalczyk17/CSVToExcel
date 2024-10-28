import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFTable;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class CSVToExcelTable {

    public static void CSVToExcel(String csvFilePath, String excelFilePath) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Formatted Table");

            // Read CSV data
            List<String> lines = Files.readAllLines(Paths.get(csvFilePath));
            int rowNum = 0;

            // Create cell styles for centering, number formatting, and percentage formatting
            CellStyle centerStyle = workbook.createCellStyle();
            centerStyle.setAlignment(HorizontalAlignment.CENTER);
            centerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            CellStyle numberStyle = workbook.createCellStyle();
            DataFormat format = workbook.createDataFormat();
            numberStyle.setDataFormat(format.getFormat("0.00")); // Adjust the format as needed

            CellStyle percentageStyle = workbook.createCellStyle();
            percentageStyle.setDataFormat(format.getFormat("0.00%")); // Percentage format

            // Create the first header row (for group labels like Attack, Service)
            Row groupHeaderRow = sheet.createRow(rowNum++);

            // Write CSV rows to Excel sheet starting from the second row
            Row headerRow = sheet.createRow(rowNum++);
            String[] values = lines.get(0).split(",");

            // Define mapping of old headers to new headers
            Map<String, String> headerMapping = new HashMap<>();
            headerMapping.put("Attack K", "Punkty");
            headerMapping.put("Attack E", "Błędy");
            headerMapping.put("Attack Atk%", "% ataków zakończonych punktem");
            headerMapping.put("Attack Atk% Trans", "Atk% Kontra");
            headerMapping.put("Attack Kill%", "Punkt%");
            headerMapping.put("Attack K/S", "Punkt/Set");
            headerMapping.put("Attack Error%", "Błąd%");
            headerMapping.put("Serve SA", "Asy");
            headerMapping.put("Serve SE", "Serwis błąd");
            headerMapping.put("Serve TA", "Serwisy");
            headerMapping.put("Receive TA", "Przyjęcia");
            headerMapping.put("Receive Pass%", "Średni % przyjęcia");
            headerMapping.put("Perfect Perfect Pass%", "% Idealnego przyjęcia");
            headerMapping.put("Dig DS", "Udane obrony");
            headerMapping.put("Dig DE", "Błąd obrony");
            headerMapping.put("Block BS", "Punkty blokiem");
            headerMapping.put("Block BE", "Błąd w bloku");
            headerMapping.put("Block B/S", "Blok/Set");

            // Write headers to headerRow with modifications
            for (int i = 0; i < values.length; i++) {
                String header = values[i].trim();
                String newHeader = headerMapping.getOrDefault(header, header); // Get new header or keep original
                headerRow.createCell(i).setCellValue(newHeader);

                // Assign each column to a category group
                if (header.startsWith("ATK") || header.contains("Attack")) {
                    groupHeaderRow.createCell(i).setCellValue("Atak");
                } else if (header.startsWith("Serv")) {
                    groupHeaderRow.createCell(i).setCellValue("Serwis");
                } else if (header.startsWith("Rec")) {
                    groupHeaderRow.createCell(i).setCellValue("Przyjęcie");
                } else if (header.startsWith("Blk") || header.contains("Block")) {
                    groupHeaderRow.createCell(i).setCellValue("Obrona");
                } else if (header.startsWith("Dig") || header.contains("Dig")) {
                    groupHeaderRow.createCell(i).setCellValue("Blok"); // No category
                } else {
                    groupHeaderRow.createCell(i).setCellValue(""); // No category
                }
            }

            // Merge cells for group headers
            mergeCategoryCells(sheet, groupHeaderRow, "Atak");
            mergeCategoryCells(sheet, groupHeaderRow, "Serwis");
            mergeCategoryCells(sheet, groupHeaderRow, "Przyjęcie");
            mergeCategoryCells(sheet, groupHeaderRow, "Obrona");
            mergeCategoryCells(sheet, groupHeaderRow, "Blok");

            // Populate remaining rows with data
            for (int j = 1; j < lines.size(); j++) {
                Row row = sheet.createRow(rowNum++);
                String[] rowData = lines.get(j).split(",");
                for (int i = 0; i < rowData.length; i++) {
                    Cell cell = row.createCell(i);
                    try {
                        double value = Double.parseDouble(rowData[i].trim()); // Assuming numeric values
                        cell.setCellValue(value); // Set numeric value

                        // Apply styles based on header
                        String header = headerRow.getCell(i).getStringCellValue();
                        if (header.endsWith("%") || header.contains("%")) {
                            cell.setCellStyle(percentageStyle); // Set percentage style
                        } else {
                            cell.setCellStyle(numberStyle); // Set number style
                        }
                    } catch (NumberFormatException e) {
                        cell.setCellValue(rowData[i].trim()); // Set as string if not numeric
                    }
                }
                if (rowNum == 13) { // A13
                    row.getCell(0).setCellValue("KS Mogielanka Mogielnica");
                } else if (rowNum == 14) { // A14
                    row.getCell(0).setCellValue("Przeciwnik");
                }

            }

            // Center the cells in the specified range B3:U14
            for (int r = 2; r < 14; r++) { // Rows 3 to 14 (0-indexed)
                for (int c = 1; c <= 20; c++) { // Columns B to U (1-indexed)
                    Cell cell = sheet.getRow(r).getCell(c);
                    if (cell != null) {
                        cell.setCellStyle(centerStyle);
                    }
                }
            }

            // Apply the center style to group header cells
            for (int i = 0; i < groupHeaderRow.getLastCellNum(); i++) {
                Cell cell = groupHeaderRow.getCell(i);
                cell.setCellStyle(centerStyle);
            }

            // Define the range of the table
            int numRows = sheet.getLastRowNum();
            int numCols = headerRow.getLastCellNum();
            AreaReference tableArea = new AreaReference(
                    new CellReference(1, 0),
                    new CellReference(numRows, numCols - 1),
                    workbook.getSpreadsheetVersion()
            );

            // Create the table
            XSSFTable table = ((XSSFSheet) sheet).createTable(tableArea);
            table.setName("DataTable");
            table.setDisplayName("DataTable");

            // Resize columns for better visibility
            for (int i = 0; i < numCols; i++) {
                sheet.autoSizeColumn(i);
            }

            // Save to Excel file
            try (FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {
                workbook.write(fileOut);
            }

            System.out.println("Excel file saved as '" + excelFilePath + "' with formatted table.");
        } catch (IOException e) {
            System.err.println("Error while converting CSV to Excel: " + e.getMessage());
        }
    }

    // Helper method to merge cells based on category
    private static void mergeCategoryCells(Sheet sheet, Row groupHeaderRow, String category) {
        int startCol = -1;
        for (int i = 0; i < groupHeaderRow.getLastCellNum(); i++) {
            Cell cell = groupHeaderRow.getCell(i);
            if (cell != null && cell.getStringCellValue().equals(category)) {
                if (startCol == -1) {
                    startCol = i;
                }
            } else if (startCol != -1) {
                // Merge cells from startCol to i - 1, if there are at least 2 cells to merge
                if (i - startCol > 1) {
                    sheet.addMergedRegion(new CellRangeAddress(0, 0, startCol, i - 1));
                }
                startCol = -1;
            }
        }
        // Final check to merge the last category cells if needed
        if (startCol != -1 && groupHeaderRow.getLastCellNum() - startCol > 1) {
            sheet.addMergedRegion(new CellRangeAddress(0, 0, startCol, groupHeaderRow.getLastCellNum() - 1));
        }
    }

    public static void main(String[] args) {
        String csvFilePath = "G:\\dkowalczyk\\projekty\\Nowy folder\\Java\\Stats.csv";
        String excelFilePath = "G:\\dkowalczyk\\projekty\\Nowy folder\\Java\\StatsFormatted.xlsx";
        CSVToExcel(csvFilePath, excelFilePath);
    }
}
