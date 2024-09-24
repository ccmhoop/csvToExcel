package excel;

public class ExcelConstants {

    //First entry in the array is the Main worksheet
    public final String[] sheetNames = {"Trainee", "Gender", "AgeRange", "CurrentHighestNFQ", "EmploymentStatus"
            , "MemberCompany", "OccupationalCategory", "UnemploymentTime", "CountyOfResidence", "AttendedEvent"};

    public final String RELS = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
            """;

}