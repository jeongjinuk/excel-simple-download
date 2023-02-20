package org.example.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;

public final class NoStyle implements Style{

    @Override
    public void configure(CellStyle cellStyle) {
        // do nothing
    }
}