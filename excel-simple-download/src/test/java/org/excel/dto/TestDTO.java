package org.excel.dto;

import lombok.Builder;
import org.excel.ExcelColumn;
import org.excel.ExcelFormula;
import org.excel.ExcelStyle;
import org.excel.formula.Counter;
import org.excel.style.DefaultBodyStyle;
import org.excel.style.DefaultStyleBorder;
import org.excel.style.ThinLine;

import java.time.LocalDate;

@Builder
@ExcelFormula(formulaClass = Counter.class)
@ExcelStyle(headerStyleClass = DefaultStyleBorder.class, bodyStyleClass = DefaultBodyStyle.class)
public class TestDTO {

    @ExcelColumn(headerName = "이름")
    private String name;

    @ExcelStyle(format = "@@@")
    @ExcelColumn(headerName = "주소")
    private String address;

    private String email;

    @ExcelColumn(headerName = "나이")
    private int age;

    @ExcelColumn(headerName = "여부")
    private boolean is;

    @ExcelStyle(bodyStyleClass = ThinLine.class, format = "AM/PM h:mm:ss;@")
    @ExcelColumn(headerName = "날짜")
    private LocalDate date;
}
