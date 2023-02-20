package org.example;

import org.example.excel.formula.Formula;

import java.lang.annotation.*;


@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelFormula {
    Class<? extends Formula>[] formulaClass();
}
