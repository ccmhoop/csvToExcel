package excel;

public class ExcelWorkbookXml extends Excel {

    private final String workbookXmlHeader = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            """;

    private final String workBookXmlFooter = "\n</workbook>";

    private final String workBookXml;

    public ExcelWorkbookXml() {
        this.workBookXml = createWorkbook();
    }

    public String getWorkBookXml() {
        return workBookXml;
    }

    private String createWorkbook() {
        return workbookXmlHeader
                + workbookXmlData()
                + workBookXmlFooter;
    }

    private StringBuilder workbookXmlData() {
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

        return sheetsBuilder.append(definedNamesBuilder);
    }

    private String sheetLine(String sheetName, int id) {
        return "<sheet name=\"" + sheetName + "\" sheetId=\"" + id + "\" r:id=\"rId" + id + "\"/>\n";
    }

    private String definedNameLine(String sheetName) {
        return "<definedName name=\"" + sheetName + "\" localSheetId=\"0\">" + sheetName + "!$A$1:$A$2000</definedName>\n";
    }

}