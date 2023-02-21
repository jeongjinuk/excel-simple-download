package org.excel.style;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.excel.excel.style.Style;

public class DefaultHeaderStyle implements Style {
    @Override
    public void configure(CellStyle cellStyle) {
        cellStyle.setBorderLeft(BorderStyle.DOUBLE);
        cellStyle.setBorderRight(BorderStyle.DOUBLE);
        cellStyle.setBorderTop(BorderStyle.DOUBLE);
        cellStyle.setBorderBottom(BorderStyle.DOUBLE);
    }


}
