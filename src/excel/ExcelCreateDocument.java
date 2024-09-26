package excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class ExcelCreateDocument extends Excel {

    private final ExcelContentTypes contentTypes;
    private final ExcelWorkbookRelations workbookRelations;
    private final ExcelWorkbookXml workbookXml;
    private final ExcelWorksheets worksheets;
    private final ExcelStyles styleSheet;

    public ExcelCreateDocument(String fileName, String filePath) throws IOException {
        this.contentTypes = new ExcelContentTypes();
        this.workbookRelations = new ExcelWorkbookRelations();
        this.workbookXml = new ExcelWorkbookXml();
        this.styleSheet = new ExcelStyles();
        this.worksheets = new ExcelWorksheets(filePath);
        createZip(fileName);
    }

    private void createZip(String fileName) {
        try (FileOutputStream fos = new FileOutputStream(fileName + ".xlsx");
             ZipOutputStream zos = new ZipOutputStream(fos)) {
            // Create [Content_Types].xml
            addToZip(zos, "[Content_Types].xml", contentTypes.getContentTypes());
            // Create xl/_rels/workbook.xml.rels
            addToZip(zos, "xl/_rels/workbook.xml.rels", workbookRelations.getWorkbookRelations());
            // Create xl/workbook.xml
            addToZip(zos, "xl/workbook.xml", workbookXml.getWorkBookXml());
            // Create xl/styles.xml
            addToZip(zos, "xl/styles.xml", styleSheet.getStylesXml());
            //  Create xl/worksheets.xml
            for (Map.Entry<String, String> entry : worksheets.getExcelWorksheets().entrySet()) {
                addToZip(zos, "xl/worksheets/" + entry.getKey() + ".xml", entry.getValue());
            }
            // Create _rels/.rels
            addToZip(zos, "_rels/.rels", RELS);
            System.out.println(fileName + ".xlsx created successfully");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Helper method to add files to the ZIP
    private void addToZip(ZipOutputStream zos, String fileName, String content) throws IOException {
        zos.putNextEntry(new ZipEntry(fileName));
        zos.write(content.getBytes());
        zos.closeEntry();
    }
}
