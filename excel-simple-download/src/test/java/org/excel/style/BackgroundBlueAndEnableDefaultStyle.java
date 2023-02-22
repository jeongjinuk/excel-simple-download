package org.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.excel.excel_file.style.Style;

public class BackgroundBlueAndEnableDefaultStyle implements Style {
    @Override
    public boolean enableDefaultStyle() {
        return true;
    }

    @Override
    public void configure(CellStyle cellStyle) {
        cellStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }
}
