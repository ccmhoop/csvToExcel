package excel;

public class ExcelStyles extends Excel {

    private final String stylesHeader = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            """;

    private final String stylesFooter = "</styleSheet>";

    public String getStylesXml() {
        return stylesHeader
                + getFonts()
                + getFills()
                + getBorders()
                + getCellFormats()
                + stylesFooter;
    }

    // Define fonts (if needed, in this case, no special font settings)
    private String getFonts() {
        return """
            <fonts count="1">
                <font>
                    <sz val="11"/>
                    <color theme="1"/>
                    <name val="Calibri"/>
                    <family val="2"/>
                </font>
            </fonts>
            """;
    }

    // Define fills for both styles (Blue for A-H,L and Green for I,J,K,M)
    private String getFills() {
        return """
            <fills count="3">
                <fill>
                    <patternFill patternType="none"/>
                </fill>
                <fill>
                    <patternFill patternType="solid">
                        <fgColor rgb="FF008CCC"/> <!-- Light Blue -->
                        <bgColor indexed="64"/>
                    </patternFill>
                </fill>
                <fill>
                    <patternFill patternType="solid">
                        <fgColor rgb="FF61BF1A"/> <!-- Light Green -->
                        <bgColor indexed="64"/>
                    </patternFill>
                </fill>
            </fills>
            """;
    }

    // Define borders (if needed, in this case, no specific borders)
    private String getBorders() {
        return """
            <borders count="1">
                <border>
                    <left/>
                    <right/>
                    <top/>
                    <bottom/>
                    <diagonal/>
                </border>
            </borders>
            """;
    }

    // Define the cell formats, mapping style ids to the fills defined above
    private String getCellFormats() {
        return """
            <cellXfs count="3">
                <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
                <xf numFmtId="0" fontId="0" fillId="1" borderId="0" xfId="0"/> <!-- Blue Style -->
                <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0"/> <!-- Green Style -->
            </cellXfs>
            """;
    }
}