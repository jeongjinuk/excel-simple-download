package org.excel.excel.sheet;

import java.util.regex.MatchResult;
import java.util.regex.Pattern;
import java.util.stream.Stream;

public enum CellAddressType {
    ROW_ABS,
    COL_ABS,
    ABS;

    private static final String ABS_PATTERN = "([A-Z]+)([0-9]+)";
    public static String convert(String str, CellAddressType type){
        switch (type){
            case ABS:
                return convertABS(str);
            case COL_ABS:
                return convertColABS(str);
            case ROW_ABS:
                return convertRowABS(str);
            default:
                throw new UnsupportedOperationException();
        }
    }

    private static String convertRowABS(String str){
        return  str.replaceAll(ABS_PATTERN,"$1\\$$2");
    }
    private static String convertColABS(String str){
        return str.replaceAll(ABS_PATTERN,"\\$$1$2");
    }

    private static String convertABS(String str){
        return str.replaceAll(ABS_PATTERN,"\\$$1\\$$2");
    }
}
