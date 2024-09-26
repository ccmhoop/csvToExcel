package excel;

public class ExcelContentTypes extends Excel {

    private final String contentTypesHeader = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
            <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
            """;

    private final String contentTypesFooter = """
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            </Types>
            """;

    private final String contentTypes;

    public ExcelContentTypes() {
        this.contentTypes = createWorkbook();
    }

    public String getContentTypes() {
        return contentTypes;
    }

    private String createWorkbook() {
        return contentTypesHeader
                + contentTypesData()
                + contentTypesFooter;
    }

    private StringBuilder contentTypesData() {
        StringBuilder contentTypes = new StringBuilder();
        for (String sheetName : sheetNames) {
            contentTypes.append(overrideLine(sheetName));
        }
        return contentTypes;
    }

    private String overrideLine(String name) {
        return "<Override PartName=\"/xl/worksheets/" + name + ".xml\" "
                + "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n";
    }

}