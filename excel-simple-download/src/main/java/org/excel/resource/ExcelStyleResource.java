package org.excel.resource;

import org.apache.poi.ss.usermodel.CellStyle;
import org.excel.excel.style.StyleLocation;

import java.util.HashMap;
import java.util.Map;

public final class ExcelStyleResource {
    private final Map<String, CellStyle> headerMap = new HashMap<>();
    private final Map<String, CellStyle> bodyMap = new HashMap<>();
    ExcelStyleResource() {}
    void put(StyleLocation styleLocation, String key, CellStyle cellStyle){
        switch (styleLocation){
            case HEADER:
                headerMap.put(key,cellStyle);
                return;
            case BODY:
                bodyMap.put(key,cellStyle);
                return;
            default:
                throw new UnsupportedOperationException();
        }
    }
    public CellStyle getCellStyle(StyleLocation styleLocation, String fieldName){
        switch (styleLocation){
            case HEADER:
                return headerMap.get(fieldName);
            case BODY:
                return bodyMap.get(fieldName);
            default:
                throw new UnsupportedOperationException();
        }
    }
}