package excel;

public class ExcelWorkbookXml extends ExcelConstants {

    private final String workBookXml;

    private final String workbookXmlHeader = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            """;

    private final String workBookXmlFooter = "\n</workbook>";

    public ExcelWorkbookXml() {
        this.workBookXml = workbookXmlBuilder();
    }

    public String getWorkBookXml() {
        return workBookXml;
    }

    private String workbookXmlBuilder() {
        StringBuilder sheetsBuilder = new StringBuilder();
        StringBuilder definedNamesBuilder = new StringBuilder();

        sheetsBuilder.append("<sheets>\n");
        definedNamesBuilder.append("<definedNames>\n");

        int id = 1;
        for (String sheetName : sheetNames) {
            sheetsBuilder.append(sheetLine(sheetName, id));
            definedNamesBuilder.append(definedNameLine(sheetName));
            id++;
        }

        sheetsBuilder.append("</sheets>\n");
        definedNamesBuilder.append("</definedNames>");

        return workbookXmlHeader + sheetsBuilder + definedNamesBuilder + workBookXmlFooter;
    }

    private String sheetLine(String sheetName, int id) {
        return "<sheet name=\"" + sheetName + "\" sheetId=\"" + id + "\" r:id=\"rId" + id + "\"/>\n";
    }

    private String definedNameLine(String sheetName) {
        return "<definedName name=\"" + sheetName + "\" localSheetId=\"0\">" + sheetName + "!$A$1:$A$2000</definedName>\n";
    }

}