package org.example.excel.sheet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public interface SheetHelper {
    Row getRow(int i);
    Cell getCell(int row, int col);
    Cell getCellByFieldName(int row, String fieldName);
    String getCellAddress(int row, int col);
    String getCellAddressByFieldName(int row, String fieldName);
    String getCellAddress(int row, int col, CellAddressType type);
    String getCellAddressByFieldName(int row, String fieldName, CellAddressType type);
    boolean containsFieldName(String fieldName);
    void setCellValue(Cell cell, Object o);
    void setFormulaCell(Cell cell, String formula);
}
