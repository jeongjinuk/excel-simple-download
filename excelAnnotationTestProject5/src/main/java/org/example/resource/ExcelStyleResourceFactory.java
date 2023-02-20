package org.example.resource;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.example.ExcelStyle;
import org.example.excel.style.NoStyle;
import org.example.excel.style.Style;
import org.example.excel.style.StyleLocation;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Optional;
import java.util.function.Function;

public final class ExcelStyleResourceFactory {
    private CellStyle defaultHeaderStyle;
    private CellStyle defaultBodyStyle;
    private final Workbook workbook;
    private final List<Field> fields;
    private final Optional<ExcelStyle> defaultStyle;

    public ExcelStyleResourceFactory(List<Field> fields, Optional<ExcelStyle> defaultStyle, Workbook workbook) {
        this.workbook = workbook;
        this.fields = fields;
        this.defaultStyle = defaultStyle;
        this.defaultHeaderStyle = workbook.createCellStyle();
        this.defaultBodyStyle = workbook.createCellStyle();
    }

    public ExcelStyleResource createExcelStyleResource() {
        ExcelStyleResource excelStyleResource = new ExcelStyleResource();
        defaultStyle.ifPresent(this :: configDefaultCellStyle);
        fields.stream()
                .filter(field -> isExcelStyleAnnotationPresent(field, excelStyleResource))
                .forEach(field -> configCellStyle(field,excelStyleResource));
        return excelStyleResource;
    }

    private void configDefaultCellStyle(ExcelStyle excelStyle){
        this.defaultHeaderStyle = config(excelStyle.headerStyleClass(), cellStyle -> cellStyle, false);
        this.defaultBodyStyle = config(excelStyle.bodyStyleClass(), CellStyle -> CellStyle, false);
        setDataFormat(this.defaultBodyStyle, excelStyle.format());
    }

    private void configCellStyle(Field field, ExcelStyleResource excelStyleResource) {
        ExcelStyle annotation = field.getDeclaredAnnotation(ExcelStyle.class);

        CellStyle headerStyle = config(annotation.headerStyleClass(), CellStyle -> cloningStyle(defaultHeaderStyle, CellStyle),true);
        CellStyle bodyStyle = config(annotation.bodyStyleClass(), CellStyle -> cloningStyle(defaultBodyStyle, CellStyle), true);
        setDataFormat(bodyStyle, annotation.format());

        excelStyleResource.put(StyleLocation.HEADER,field.getName(), headerStyle);
        excelStyleResource.put(StyleLocation.BODY,field.getName(), bodyStyle);
    }

    private boolean isExcelStyleAnnotationPresent(Field field, ExcelStyleResource excelStyleResource){
        if(!field.isAnnotationPresent(ExcelStyle.class)){
            excelStyleResource.put(StyleLocation.HEADER, field.getName(), defaultHeaderStyle);
            excelStyleResource.put(StyleLocation.BODY, field.getName(), defaultBodyStyle);
            return false;
        }
        return true;
    }

    private CellStyle config(Class<? extends Style> c, Function<CellStyle,CellStyle> enableCloningFunction, boolean enableDefault) {
        CellStyle cellStyle = workbook.createCellStyle();
        if (isNoStyleClass(c)) {
            return enableCloningFunction.apply(cellStyle);
        }
        Style style = ReflectionUtils.getInstance(c);
        if (enableDefault && style.usedDefaultStyle()){
            cellStyle = enableCloningFunction.apply(cellStyle);
        }
        style.configure(cellStyle);
        return cellStyle;
    }

    private void setDataFormat(CellStyle cellStyle, String format){
        cellStyle.setDataFormat(workbook.createDataFormat().getFormat(format));
    }

    private CellStyle cloningStyle(CellStyle from, CellStyle to){
        to.cloneStyleFrom(from);
        return to;
    }

    private boolean isNoStyleClass(Class<? extends Style> c) {
        return NoStyle.class.equals(c);
    }
}
