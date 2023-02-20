package org.example;

import org.example.excel.style.NoStyle;
import org.example.excel.style.Style;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.FIELD,ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelStyle {
    Class<? extends Style> headerStyleClass() default NoStyle.class;
    Class<? extends Style> bodyStyleClass() default NoStyle.class;
    String format() default "";
}
