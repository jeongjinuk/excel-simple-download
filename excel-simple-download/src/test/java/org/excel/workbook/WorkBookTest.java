package org.excel.workbook;

import org.apache.poi.ss.usermodel.*;
import org.excel.dto.TestDTO;
import org.excel.excel.sheet.SheetHelper;
import org.excel.style.BackgroundBlueAndEnableDefaultStyle;
import org.excel.style.DefaultBodyStyle;
import org.excel.style.DefaultHeaderStyle;
import org.excel.style.ThinLine;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.time.LocalDate;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class WorkBookTest {
    private List<TestDTO> list;

    @BeforeEach
    void init() {
        String email = "@test.com";
        this.list = IntStream.rangeClosed(0, 50)
                .mapToObj(operand ->
                        TestDTO.builder()
                                .name(String.valueOf(operand * 10000))
                                .address(String.format("%s, %s%s - %s", operand, operand, operand, operand))
                                .email(operand + email)
                                .age(operand)
                                .is(operand % 2 == 1)
                                .build()
                ).collect(Collectors.toList());
    }
    @Test
    @DisplayName("엑셀 DTO 설정 정상동작 여부 확인 테스트")
    void SXXFSExcelFileTest() {
        MySXXFS<TestDTO> mySXXFS = new MySXXFS<>(list, TestDTO.class);
        Workbook workbook = mySXXFS.getWorkbook();
        assertThat(workbook);
    }
    void assertThat(Workbook workbook) {
        assertFormula(workbook);
        assertStyle(workbook);
    }
    @DisplayName("formula 정상 동작 테스트, 수식에 countif(range,'*true') 수식이 정상적으로 들어가있는지 확인")
    void assertFormula(Workbook workbook) {
        Sheet sheet = workbook.getSheetAt(0);
        String counterFormula = "COUNTIF(D2:D51,\"*true\")";
        String cellFormula = sheet.getRow(list.size()).getCell(3).getCellFormula();
        assertEquals(counterFormula, cellFormula);
    }
    @DisplayName("DTO에 설정된 style이 정상적으로 동작하는지 확인 : 확인 필드 목록 [name, address, date]")
    void assertStyle(Workbook workbook) {
        Sheet sheet = workbook.getSheetAt(0);

        List<Row> rows = new ArrayList<>();
        sheet.iterator().forEachRemaining(rows :: add);
        Map<String, List<Cell>> excelSheetMap = getExcelSheetMap(rows);

        Cell headerName = excelSheetMap.get("header").get(0);
        Cell headerDate = excelSheetMap.get("header").get(4);
        Cell bodyAddress = excelSheetMap.get("body").get(1);
        Cell bodyDate = excelSheetMap.get("body").get(4);

        Map<String, CellStyle> style = getStyle(workbook);

        assertStyle(style.get("headerName"), headerName.getCellStyle());
        assertStyle(style.get("headerDate"), headerDate.getCellStyle());
        assertStyle(style.get("bodyAddress"), bodyAddress.getCellStyle());
        assertStyle(style.get("bodyDate"), bodyDate.getCellStyle());

        assertEquals(workbook.createDataFormat().getFormat("@@@"), bodyAddress.getCellStyle().getDataFormat());
        assertEquals(workbook.createDataFormat().getFormat("AM/PM h:mm:ss;@"), bodyDate.getCellStyle().getDataFormat());
    }
    void assertStyle(CellStyle o1, CellStyle o2){
        assertEquals(o1.getBorderBottom(), o2.getBorderBottom());
        assertEquals(o1.getBorderLeft(), o2.getBorderLeft());
        assertEquals(o1.getBorderRight(), o2.getBorderRight());
        assertEquals(o1.getBorderTop(), o2.getBorderTop());
        assertEquals(o1.getFillBackgroundColor(), o2.getFillBackgroundColor());
        assertEquals(o1.getFillForegroundColor(), o2.getFillForegroundColor());
        assertEquals(o1.getFillPattern(), o2.getFillPattern());
    }
    Map<String, CellStyle> getStyle(Workbook workbook) {
        Map<String, CellStyle> map = new HashMap<>();

        DefaultHeaderStyle defaultHeaderStyle = new DefaultHeaderStyle();
        DefaultBodyStyle defaultBodyStyle = new DefaultBodyStyle();
        BackgroundBlueAndEnableDefaultStyle backgroundBlueAndEnableDefaultStyle = new BackgroundBlueAndEnableDefaultStyle();
        ThinLine thinLine = new ThinLine();

        CellStyle headerName = workbook.createCellStyle();
        CellStyle headerDate = workbook.createCellStyle();
        CellStyle bodyAddress = workbook.createCellStyle();
        CellStyle bodyDate = workbook.createCellStyle();

        short addressFormat = workbook.createDataFormat().getFormat("@@@");
        short dateFormat = workbook.createDataFormat().getFormat("AM/PM h:mm:ss;@");

        bodyAddress.setDataFormat(addressFormat);
        bodyDate.setDataFormat(dateFormat);

        defaultHeaderStyle.configure(headerName);
        defaultBodyStyle.configure(bodyAddress);
        backgroundBlueAndEnableDefaultStyle.configure(bodyAddress);
        thinLine.configure(headerDate);
        defaultBodyStyle.configure(bodyDate);


        map.put("headerName", headerName);
        map.put("headerDate", headerDate);
        map.put("bodyAddress", bodyAddress);
        map.put("bodyDate", bodyDate);
        return map;
    }
    Map<String, List<Cell>> getExcelSheetMap(List<Row> rows) {
        Map<String, List<Cell>> map = new HashMap<>();
        Row header = rows.remove(0);
        Row body = rows.remove(1);

        List<Cell> headerCells = new ArrayList<>();
        List<Cell> bodyCells = new ArrayList<>();

        header.iterator().forEachRemaining(headerCells :: add);
        body.iterator().forEachRemaining(bodyCells :: add);

        map.put("header", headerCells);
        map.put("body", bodyCells);
        return map;
    }
}
