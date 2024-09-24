package excel;

public class ExcelContentTypes extends ExcelConstants {

    private final String contentTypes;

    private final String contentTypesHeader = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
            """;

    private final String contentTypesFooter = """
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            </Types>
            """;

    public ExcelContentTypes() {
        this.contentTypes = contentTypesBuilder();
    }

    public String getContentTypes() {
        return contentTypes;
    }

    private String contentTypesBuilder() {
        StringBuilder contentTypes = new StringBuilder();
        for (String sheetName : sheetNames) {
            contentTypes.append(overrideLine(sheetName));
        }
        return contentTypesHeader + contentTypes + contentTypesFooter;
    }

    private String overrideLine(String name) {
        return "<Override PartName=\"/xl/worksheets/" + name + ".xml\" "
                + "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n";
    }

}