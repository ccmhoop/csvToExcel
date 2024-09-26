package excel;

import xml.XmlUtility;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelDataValidations {

    private final StringBuilder validationBlockBuilder;
    private final Map<String, String> sheetData;
    private final int maxRowRange;
    private String dataValidations;
    private int worksheetCount = 0;

    public ExcelDataValidations(int maxRowRange) throws IOException {
        this.maxRowRange = maxRowRange;
        this.sheetData = new HashMap<>();
        this.validationBlockBuilder = new StringBuilder();
        processFilesInFolder("src/excel/droplistRecords");
    }

    public Map<String, String> getSheetData() {
        return sheetData;
    }

    public String getDataValidations() {
        return dataValidations;
    }

    private void processFilesInFolder(String folderPath) throws IOException {
        File dir = new File(folderPath);
        File[] directoryListing = dir.listFiles();
        String filename;
        for (File file : directoryListing) {
            filename = file.getName().replaceAll(".txt", "");
            createSheetData(filename, file);
            validationBlockBuilder.append(dataValidationBlock(filename));
        }

        dataValidations = createDataValidations();
    }

    private void createSheetData(String filename, File file) throws IOException {
        List<String> formattedRecords = readAndFormatRecords(file);
        //map key = filename without File extension : example.txt -> example;
        String sheetData = buildSheetData(formattedRecords);
        this.sheetData.put(filename, sheetData);
        worksheetCount++;
    }

    private List<String> readAndFormatRecords(File file) throws IOException {
        List<String> formattedRecords = new ArrayList<>();
        try (BufferedReader br = new BufferedReader(new FileReader(file))) {
            String record;
            while ((record = br.readLine()) != null) {
                if (!record.isEmpty()) {
                    formattedRecords.add(XmlUtility.replaceSpecialCharacters(record));
                }
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return formattedRecords;
    }

    private String buildSheetData(List<String> formattedRecords) {
        StringBuilder sheetDataLineBuilder = new StringBuilder();
        int rowCounter = 1;
        for (String record : formattedRecords) {
            sheetDataLineBuilder.append(sheetDataLine(record, rowCounter));
            rowCounter++;
        }
        return sheetDataLineBuilder.toString();
    }

    private String sheetDataLine(String record, int rowCounter) {
        return "<row r=\"" + rowCounter + "\">"
                + "<c r=\"A" + rowCounter + "\" t=\"inlineStr\"><is><t>"
                + record
                + "</t></is></c></row>\n";
    }

    private String dataValidationBlock(String fileName) {
        char columnChar = getColumnChar(fileName);
        return "<dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" "
                + "sqref=\"" + columnChar + "2:" + columnChar + maxRowRange + "\">\n"
                + "<formula1>=" + fileName + "</formula1>\n"
                + "</dataValidation>\n";
    }

    private String createDataValidations() {
        return "<dataValidations count=\"" + worksheetCount + "\">\n"
                + validationBlockBuilder
                + "</dataValidations>\n";
    }

    private char getColumnChar(String filename) {
        return switch (filename) {
            case "Gender" -> 'C';
            case "AgeRange" -> 'D';
            case "CurrentHighestNFQ" -> 'G';
            case "EmploymentStatus" -> 'H';
            case "MemberCompany" -> 'I';
            case "OccupationalCategory" -> 'J';
            case "UnemploymentTime" -> 'k';
            case "CountyOfResidence" -> 'L';
            case "AttendedEvent" -> 'M';
            default -> '0';
        };
    }
}