package org.example.dto;

import lombok.Builder;
import org.example.ExcelColumn;
import org.example.ExcelFormula;
import org.example.ExcelStyle;
import org.example.formula.Counter;
import org.example.style.DefaultBodyStyle;
import org.example.style.DefaultStyleBorder;
import org.example.style.ThinLine;

import java.time.LocalDate;
import java.time.LocalTime;

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
