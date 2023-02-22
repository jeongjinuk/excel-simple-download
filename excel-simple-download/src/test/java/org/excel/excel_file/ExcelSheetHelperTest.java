package org.excel.excel_file;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.excel.exception.NoSuchFieldNameException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.HashMap;

import static org.junit.jupiter.api.Assertions.*;

class ExcelSheetHelperTest {
    @Test
    @DisplayName("excelSheetHelper에서 명시적으로 표시한 예외가 정상작동하는지 여부 테스트, 필드이름을 가지고 cell 의 주소를 반환하는 메서드들")
    void excelSheetHelperExceptionTest(){
        ExcelSheetHelper excelSheetHelper = new ExcelSheetHelper(new HashMap<>());

        assertThrows(NoSuchFieldNameException.class, () -> excelSheetHelper.getCellByFieldName(0, "nothing"));
        assertThrows(NoSuchFieldNameException.class, () -> excelSheetHelper.getCellAddressByFieldName(0, "nothing"));
    }

    @Test
    @DisplayName("excelSheetHelper의 최상단 기능 테스트, getCell(), getRow(), getCellAddress(), containsFieldName() 이 기능들을 다른 기능에서 사용하기 때문에 이것만 테스트")
    void excelSheetHelperFunctionTest(){
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFRow row = workbook.createSheet().createRow(0);
        XSSFCell cell = row.createCell(0);

        cell.setCellValue("test");

        XSSFSheet sheet = workbook.getSheetAt(0);
        ExcelSheetHelper excelSheetHelper = new ExcelSheetHelper(new HashMap<>());
        excelSheetHelper.setSheet(sheet);

        assertThat(row, cell, excelSheetHelper);
    }

    void assertThat(Row row, Cell cell, ExcelSheetHelper excelSheetHelper){
        assertEquals(row, excelSheetHelper.getRow(0));
        assertEquals(cell, excelSheetHelper.getCell(0,0));
        assertEquals(cell.getAddress().formatAsString(), excelSheetHelper.getCellAddress(0,0));
        assertFalse(excelSheetHelper.containsFieldName("nothing"));
    }
}