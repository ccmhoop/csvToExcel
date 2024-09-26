package excel;

public class ExcelWorkbookRelations extends Excel {

    private final String workbookRelationsHeader = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            """;

    private final String workbookRelationsFooter = "</Relationships>";

    private final String workbookRelations;

    public ExcelWorkbookRelations() {
        this.workbookRelations = createWorkbook();
    }

    public String getWorkbookRelations() {
        return workbookRelations;
    }

    private String createWorkbook() {
        return workbookRelationsHeader
                + workbookRelationsData()
                + workbookRelationsFooter;
    }

    private StringBuilder workbookRelationsData() {
        StringBuilder relationshipLineBuilder = new StringBuilder();
        relationshipLineBuilder.append(styleRelationshipLine());
        int id = 1;
        for (String sheetName : sheetNames) {
            relationshipLineBuilder.append(relationshipLine(sheetName, id));
            id++;
        }
        return relationshipLineBuilder;
    }

    private String styleRelationshipLine() {
        return "<Relationship Id=\"rId" + (sheetNames.length + 1) + "\" "
                + "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" "
                + "Target=\"styles.xml\"/>\n";
    }

    private String relationshipLine(String name, int id) {
        return "<Relationship Id=\"rId" + id + "\" "
                + "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
                + "Target=\"worksheets/" + name + ".xml\"/>\n";
    }

}