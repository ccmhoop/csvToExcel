package excel;

public class ExcelWorkbookRelations extends ExcelConstants {

    private final String workbookRelations;

    private final String workbookRelationsHeader = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            """;

    private final String workbookRelationsFooter = "</Relationships>";

    public ExcelWorkbookRelations() {
        this.workbookRelations = workbookRelationsBuilder();
    }

    public String getWorkbookRelations() {
        return workbookRelations;
    }

    private String workbookRelationsBuilder() {
        StringBuilder relationshipLineBuilder = new StringBuilder();
        int id = 1;
        for (String sheetName : sheetNames) {
            relationshipLineBuilder.append(relationshipLine(sheetName, id));
            id++;
        }
        return workbookRelationsHeader + relationshipLineBuilder + workbookRelationsFooter;
    }

    private String relationshipLine(String name, int id) {
        return "<Relationship Id=\"rId" + id + "\" "
                + "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
                + "Target=\"worksheets/" + name + ".xml\"/>\n";
    }

}