package org.example.excel.formula;

import org.example.excel.sheet.SheetHelper;

import java.util.List;
public interface Formula {
    default void renderFormula(SheetHelper sheet, int criteriaRowIndex, List<?> data){
        validate(sheet, criteriaRowIndex, data);
        render(sheet, criteriaRowIndex, data);
    }
    void validate(SheetHelper sheet, int criteriaRowIndex, List<?> data);
    void render(SheetHelper sheet, int criteriaRowIndex, List<?> data);
}