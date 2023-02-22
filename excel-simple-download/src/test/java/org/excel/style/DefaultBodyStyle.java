package org.excel.style;

import org.apache.poi.ss.usermodel.*;
import org.excel.excel_file.style.Style;

public class DefaultBodyStyle implements Style {
    @Override
    public void configure(CellStyle cellStyle) {
        cellStyle.setBorderLeft(BorderStyle.MEDIUM);
        cellStyle.setBorderRight(BorderStyle.MEDIUM);
        cellStyle.setBorderTop(BorderStyle.MEDIUM);
        cellStyle.setBorderBottom(BorderStyle.MEDIUM);
    }
}
