package excel;

import csv.CsvReader;
import xml.XmlUtility;

import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelWorksheets extends Excel {

    private final List<List<Object>> csvAsList;
    private final ExcelDataValidations dataValidations;
    private final Map<String, String> excelWorksheets;
    private final int rowAmount;

    private final List<Object> columnLabels = List.of("FirstName*", "Surname*", "Gender*", "AgeRange*"
            , "EmailAddress*", "Phone*", "CurrentHighestNFQ*", "EmploymentStatus*", "MemberCompany"
            , "OccupationalCategory", "UnemploymentTime", "CountyOfResidence*", "AttendedEvent");

    private final String worksheetHeader = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            """;

    private final String worksheetFooter = "</worksheet>";

    public ExcelWorksheets(String filePath) throws IOException {
        CsvReader<Object> csvReader = new CsvReader<>(filePath);
        this.csvAsList = csvReader.getRecordsAsList();
        this.rowAmount = csvAsList.size();
        this.dataValidations = new ExcelDataValidations(rowAmount + 20);
        this.excelWorksheets = new HashMap<>();
        createWorkbook();
    }

    public Map<String, String> getExcelWorksheets() {
        return excelWorksheets;
    }

    private void createWorkbook() {
        csvAsList.set(0, columnLabels);
        for (int i = 0; i < sheetNames.length; i++) {
            excelWorksheets.put(sheetNames[i], buildWorksheet(i));
        }
    }

    private String buildWorksheet(int index) {
        return worksheetHeader
                + "<sheetData>\n"
                + (index == 0 ? mainSheetData() : droplistSheetData(index))
                + "</sheetData>\n"
                + (index == 0 ? dataValidations.getDataValidations() : "")
                + worksheetFooter;
    }

    private StringBuilder mainSheetData() {
        StringBuilder rowRecordsBuilder = new StringBuilder();
        int row = 1;
        for (List<Object> rowRecords : csvAsList) {
            char columnLetter = 'A';
            rowRecordsBuilder.append("<row r=\"").append(row).append("\">\n");
            for (Object record : rowRecords) {
                if (row == 1){
                    rowRecordsBuilder.append(sheetDataLineWithStyle(columnLetter, row, record));
                }else{
                    rowRecordsBuilder.append(sheetDataLine(columnLetter, row, record));
                }

                columnLetter++;
            }
            rowRecordsBuilder.append("</row>\n");
            row++;
        }
        return rowRecordsBuilder;
    }

    private String droplistSheetData(int index) {
        return dataValidations.getSheetData().get(sheetNames[index]);
    }

    private String sheetDataLine(char columnLetter, int row, Object record) {
        record = XmlUtility.replaceSpecialCharacters(record.toString());
        return "<c t=\"inlineStr\" r=\"" + columnLetter + row + "\"><is><t>" + record + "</t></is></c>\n";
    }

    private String sheetDataLineWithStyle(char columnLetter, int row, Object record) {
        record = XmlUtility.replaceSpecialCharacters(record.toString());
        String style = getColumnStyle(columnLetter);
        return "<c t=\"inlineStr\" r=\"" + columnLetter + row + "\" s=\"" + style + "\"><is><t>" + record + "</t></is></c>\n";
    }

    private String getColumnStyle (char columnLetter){
        return switch (columnLetter){
            case 'I','J','K','M' -> "2";
            default -> "1";
        };
    }

}