package org.excel.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;

public interface Style {
    default boolean enableDefaultStyle(){
        return false;
    }
    void configure(CellStyle cellStyle);
}