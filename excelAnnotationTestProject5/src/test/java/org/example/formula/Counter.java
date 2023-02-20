package org.example.formula;

import org.apache.poi.ss.usermodel.Cell;
import org.example.excel.formula.Formula;
import org.example.excel.sheet.SheetHelper;

import java.util.List;

public class Counter implements Formula {
    /**
     * 여부 필드에서 true의 개수를 표시해주는 수식 넣어주는 역할
     * 여부 필드 하단에  =COUNTIF(RANGE:RANGE,"*true");
     */
    @Override
    public void validate(SheetHelper sheet, int criteriaRowIndex, List<?> data) {}

    @Override
    public void render(SheetHelper sheet, int criteriaRowIndex, List<?> data) {
        String formula = "COUNTIF(%s:%s,%s)";
        String criteria = "\"*true\"";
        String start = sheet.getCellAddressByFieldName(criteriaRowIndex, "is");
        String end = sheet.getCellAddressByFieldName(data.size()-1, "is");

        Cell targetCell = sheet.getCellByFieldName(data.size(), "is");

        sheet.setFormulaCell(targetCell,String.format(formula, start, end, criteria));
    }
}
