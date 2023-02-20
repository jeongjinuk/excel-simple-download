package org.excel.excel;

import org.apache.poi.ss.usermodel.*;
import org.excel.excel.style.StyleLocation;
import org.excel.resource.ExcelResource;
import org.excel.resource.ExcelResourceFactory;

import java.util.List;

public abstract class ExcelFile<T> {
    private ExcelResource excelResource;
    private Workbook workbook;
    private ExcelSheetHelper excelSheetHelper;
    protected ExcelFile(List<T> data, Class<?> clazz, Workbook workbook) {
        validate(data);
        this.workbook = workbook;
        this.excelResource = ExcelResourceFactory.prepareExcelResource(clazz, workbook);
        this.excelSheetHelper = new ExcelSheetHelper(excelResource.getFieldNameWithColumnIndexResource());
        renderExcel(data);
    }

    public Workbook getWorkbook() {
        return this.workbook;
    }

    protected void validate(List<T> data) {
    }

    protected abstract void renderExcel(List<T> data);

    protected void renderHeader(Sheet sheet, int rowIndex) {
        excelSheetHelper.setSheet(sheet);
        excelResource.getFieldResource().forEach((field, excelColumn) -> {
            Cell cell = excelSheetHelper.getCellByFieldName(rowIndex, field.getName());
            renderCellStyle(StyleLocation.HEADER, cell, field.getName());
            excelSheetHelper.setCellValue(cell, excelColumn.headerName());
        });
    }

    protected void renderBody(Sheet sheet, int rowIndex, T data) {
        excelSheetHelper.setSheet(sheet);
        excelResource.getFieldResource().forEach((field, excelColumn) -> {
            Cell cell = excelSheetHelper.getCellByFieldName(rowIndex, field.getName());
            renderCellStyle(StyleLocation.BODY, cell, field.getName());
            excelSheetHelper.setCellValue(field, cell, data);
        });
    }

    protected void renderFormula(Sheet sheet, int criteriaRowIndex, List<T> data) {
        excelSheetHelper.setSheet(sheet);
        excelResource.getFormulaResource()
                .forEach(formula -> formula.renderFormula(excelSheetHelper, criteriaRowIndex, data));
    }

    private void renderCellStyle(StyleLocation styleLocation, Cell cell, String fieldName){
        CellStyle cellStyle = excelResource.getExcelStyleResource().getCellStyle(styleLocation, fieldName);
        cell.setCellStyle(cellStyle);
    }
}
