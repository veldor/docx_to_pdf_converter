package net.veldor.docx_to_pdf.utils;

public class StringsHandler {


    public static final String EXECUTION_AREA_START = "Область исследования:";
    public static final int EXECUTION_AREA_PREFIX_LENGTH = 21;

    public static final String EXECUTION_NUMBER_START = "Номер исследования:";
    public static final int EXECUTION_NUMBER_PREFIX_LENGTH = 19;

    public static final String ALTER_EXECUTION_NUMBER_START = "ID исследования:";
    public static final int ALTER_EXECUTION_NUMBER_PREFIX_LENGTH = 17;

    public static String superTrim(String s) {
        return s.replaceAll("\\s", "");
    }

    public static String getStringFrom(String s, int prefixLength) {
        return trim(s.substring(prefixLength));
    }

    private static String trim(String s) {
        return s.trim();
    }

    public static String getExecutionArea(String value) {
        // разобью текст по переносам строк
        String[] mStrings = value.split("\n");
        for (String s : mStrings) {
            if (s.length() == 0) {
                continue;
            }
            if (s.trim().startsWith(EXECUTION_AREA_START)) {
                return StringsHandler.superTrim(StringsHandler.getStringFrom(s, EXECUTION_AREA_PREFIX_LENGTH));
            }
        }
        return null;
    }

    public static String getExecutionNumber(String value) {
        // разобью текст по переносам строк
        String[] mStrings = value.split("\n");
        for (String s : mStrings) {
            if (s.length() == 0) {
                continue;
            }
            if (s.trim().startsWith(EXECUTION_NUMBER_START)) {
                return StringsHandler.superTrim(StringsHandler.getStringFrom(s, EXECUTION_NUMBER_PREFIX_LENGTH).toUpperCase().replace("А", "A"));
            } else if (s.trim().startsWith(ALTER_EXECUTION_NUMBER_START)) {
                return StringsHandler.superTrim(StringsHandler.getStringFrom(s, ALTER_EXECUTION_NUMBER_PREFIX_LENGTH).toUpperCase().replace("А", "A"));

            }
        }
        return null;
    }
}
