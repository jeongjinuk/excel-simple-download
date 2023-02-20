package org.excel.style;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.excel.excel.style.Style;

public class ThinLine implements Style {
    @Override
    public boolean usedDefaultStyle() {
        return false;
    }

    @Override
    public void configure(CellStyle cellStyle) {
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
    }
}
