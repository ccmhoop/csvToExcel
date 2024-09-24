package xml;

public class XmlUtility {

    public static String replaceSpecialCharacters(String input) {
        input = input.replaceAll("&", "&amp;")
                .replaceAll("<", "&lt;")
                .replaceAll(">", "&gt;")
                .replaceAll("\"", "&quot;")
                .replaceAll("'", "&apos;");
        return input.trim();
    }

}
