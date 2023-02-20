package org.example.excel.sheet;

import java.util.regex.Pattern;

public enum CellAddressType {
    ROW_ABS,
    COL_ABS,
    ABS;

    private static final Pattern ROW_ABS_PATTERN = Pattern.compile("[0-9]+");
    private static final Pattern COL_ABS_PATTERN = Pattern.compile("[A-Z]+");
    public static String convert(String str, CellAddressType type){
        switch (type){
            case ABS:
                return convertABS(str).toString();
            case COL_ABS:
                return convertColABS(str).toString();
            case ROW_ABS:
                return convertRowABS(str).toString();
            default:
                throw new UnsupportedOperationException();
        }
    }

    private static StringBuilder convertRowABS(String str){
        StringBuilder stringBuilder = new StringBuilder("$");
        ROW_ABS_PATTERN.matcher(str).results()
                .forEach(stringBuilder::append);
        return stringBuilder;
    }
    private static StringBuilder convertColABS(String str){
        StringBuilder stringBuilder = new StringBuilder("$");
        COL_ABS_PATTERN.matcher(str).results()
                .forEach(stringBuilder::append);
        return stringBuilder;
    }

    private static StringBuilder convertABS(String str){
        return convertColABS(str).append(convertRowABS(str));
    }
}
