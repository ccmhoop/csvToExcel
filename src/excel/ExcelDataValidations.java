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

    private Map<String, String> droplistSheetData;
    private Map<String, String> dataValidation;
    private int droplistCount = 0;
    private final int maxRowRange;

    public ExcelDataValidations(int maxRowRange) throws IOException {
        this.maxRowRange = maxRowRange;
        createDroplistData();
    }

    public Map<String, String> getDroplistSheetData() {
        return droplistSheetData;
    }

    public String getDataValidations() {
        StringBuilder builder = new StringBuilder();
        builder.append("<dataValidations count=\"").append(droplistCount).append("\">\n");
        for (Map.Entry<String, String> entry : dataValidation.entrySet()) {
            builder.append(entry.getValue());
        }
        builder.append("</dataValidations>\n");
        return builder.toString();
    }

    private void createDroplistData() throws IOException {
        File dir = new File("src/excel/droplistRecords");
        File[] directoryListing = dir.listFiles();

        droplistSheetData = new HashMap<>();
        dataValidation = new HashMap<>();

        for (File file : directoryListing) {
            processTxtFiles(file);
        }
    }

    private void processTxtFiles(File file) {
        try (BufferedReader br = new BufferedReader(new FileReader(file))) {
            List<String> formattedRecords = readAndFormatTxtLines(br);
            String filename = file.getName().replaceAll(".txt", "");

            //map key = filename without File extension : example.txt -> example;
            droplistSheetData.put(filename, droplistSheetDataBuilder(formattedRecords));
            dataValidation.put(filename, dataValidationBuilder(filename));

            droplistCount++;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private List<String> readAndFormatTxtLines(BufferedReader br) throws IOException {
        List<String> formattedRecords = new ArrayList<>();
        String record;
        while ((record = br.readLine()) != null) {
            if (!record.isEmpty()) {
                formattedRecords.add(XmlUtility.replaceSpecialCharacters(record));
            }
        }
        return formattedRecords;
    }

    private String droplistSheetDataBuilder(List<String> formattedRecords) {
        StringBuilder sheetDataLineBuilder = new StringBuilder();
        int counter = 1;
        for (String record : formattedRecords) {
            sheetDataLineBuilder.append(sheetDataLine(record, counter));
            counter++;
        }
        return sheetDataLineBuilder.toString();
    }

    private String sheetDataLine(String record, int counter){
        return "<row r=\""+ counter +"\"><c r=\"A"+ counter +"\" t=\"inlineStr\"><is><t>"+ record +"</t></is></c></row>\n";
    }

    private String dataValidationBuilder(String fileName) {
        char columnChar = getColumnChar(fileName);
        return "<dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" " +
                "sqref=\"" + columnChar + "2:" + columnChar + maxRowRange + "\">\n" +
                "<formula1>=" + fileName + "</formula1>\n" +
                "</dataValidation>\n";
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