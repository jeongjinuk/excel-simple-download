package org.excel.workbook;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.excel.excel_file.ExcelFile;

import java.util.List;
import java.util.stream.IntStream;

public class MySXXFS<T> extends ExcelFile<T> {

    public MySXXFS(List<T> data, Class<?> clazz) {
        super(data, clazz, new SXSSFWorkbook());
    }

    @Override
    protected void validate(List<T> data) {
        SpreadsheetVersion excel2007 = SpreadsheetVersion.EXCEL2007;
        if (excel2007.getMaxRows() < data.size()){
            throw  new IllegalArgumentException(String.format("%s Excel does not support over %s rows", excel2007, excel2007.getMaxRows()));
        }
    }
    @Override
    protected void renderExcel(List<T> data) {
        Sheet sheet = getWorkbook().createSheet();
        renderHeader(sheet,0);

        IntStream.rangeClosed(1, data.size() - 1)
                .forEach(i -> renderBody(sheet, i, data.get(i)));

        renderFormula(sheet, 1, data);
    }
}
