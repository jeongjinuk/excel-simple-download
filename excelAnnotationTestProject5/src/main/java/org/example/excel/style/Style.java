package org.example.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;

public interface Style {
    default boolean usedDefaultStyle(){
        return false;
    }
    void configure(CellStyle cellStyle);
}