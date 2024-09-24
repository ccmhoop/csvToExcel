package excel;

import csv.CsvReader;
import xml.XmlUtility;

import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelWorksheets extends ExcelConstants {

    private final List<List<Object>> csvAsList;
    private final ExcelDataValidations dataValidations;
    private final Map<String, String> excelWorksheets;
    private final int rowAmount;

    private final List<Object> columnLabels = List.of("FirstName*", "Surname*", "Gender*", "AgeRange*"
            , "EmailAddress*", "Phone*", "CurrentHighestNFQ*", "EmploymentStatus*", "MemberCompany"
            , "OccupationalCategory", "UnemploymentTime", "CountyOfResidence*", "AttendedEvent");

    private final String sheetHeader = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            """;

    private final String sheetFooter = "</worksheet>";

    public ExcelWorksheets(String filePath) throws IOException {
        CsvReader<Object> csvReader = new CsvReader<>(filePath);
        this.csvAsList = csvReader.getRecordsAsList();
        this.rowAmount = csvAsList.size();
        this.dataValidations = new ExcelDataValidations(rowAmount + 20);
        this.excelWorksheets = new HashMap<>();
        sheetProcessor();
    }

    public Map<String, String> getExcelWorksheets() {
        return excelWorksheets;
    }

    private void sheetProcessor() {
        csvAsList.set(0, columnLabels);
        mainWorksheetBuilder();
        droplistWorksheetBuilder();
    }

    private void mainWorksheetBuilder() {
        StringBuilder sheetBuilder = new StringBuilder();
        int row = 1;

        sheetBuilder.append(sheetHeader);
        sheetBuilder.append("<sheetData>\n");
        for (List<Object> rowRecords : csvAsList) {
            char columnLetter = 'A';
            sheetBuilder.append("<row r=\"").append(row).append("\">\n");
            for (Object record : rowRecords) {
                String formattedRecord = XmlUtility.replaceSpecialCharacters(record.toString());
                sheetBuilder.append(sheetDataLine(columnLetter, row, formattedRecord));
                columnLetter++;
            }
            sheetBuilder.append("</row>\n");
            row++;
        }
        sheetBuilder.append("</sheetData>\n")
                .append(dataValidations.getDataValidations())
                .append(sheetFooter);

        excelWorksheets.put(sheetNames[0], sheetBuilder.toString());
    }

    private String sheetDataLine(char columnLetter, int row, String formattedRecord) {
        return "<c t=\"inlineStr\" r=\"" + columnLetter + row + "\"><is><t>" + formattedRecord + "</t></is></c>\n";
    }

    private void droplistWorksheetBuilder() {
        for (String sheetName : sheetNames) {
            if (!sheetName.equalsIgnoreCase(sheetNames[0])) {
                String droplistSheet = sheetHeader + droplistSheetData(sheetName) + sheetFooter;
                excelWorksheets.put(sheetName, droplistSheet);
            }
        }
    }

    private String droplistSheetData(String sheetName) {
        return "<sheetData>\n" + dataValidations.getDroplistSheetData().get(sheetName) + "</sheetData>\n";
    }

}