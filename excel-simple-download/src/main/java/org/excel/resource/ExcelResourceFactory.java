package org.excel.resource;

import org.apache.poi.ss.usermodel.Workbook;
import org.excel.ExcelColumn;
import org.excel.ExcelFormula;
import org.excel.ExcelStyle;
import org.excel.excel.formula.Formula;
import org.excel.exception.NotFoundExcelColumnAnnotationException;

import java.lang.reflect.Field;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

import static java.util.stream.Collectors.toMap;

public final class ExcelResourceFactory {
    private ExcelResourceFactory() {}
    public static ExcelResource prepareExcelResource(Class<?> type, Workbook workbook) {
        Map<Field, ExcelColumn> fieldResource = prepareFieldResource(type);

        validate(fieldResource, type);

        List<? extends Formula> formulaResource = prepareFormulaResource(type);

        Map<String, Integer> fieldNameWithColumnIndexResource = prepareFieldNameWithColumnIndexResource(fieldResource);

        ExcelStyleResource excelStyleResource = prepareStyleResource(type, workbook);

        return new ExcelResource(fieldResource,
                formulaResource,
                fieldNameWithColumnIndexResource,
                excelStyleResource);
    }
    private static Map<Field, ExcelColumn> prepareFieldResource(Class<?> type) {
        return ReflectionUtils.getFieldWithAnnotationList(type, ExcelColumn.class, true).stream()
                .collect(
                        toMap(
                                field -> field,
                                field -> field.getAnnotation(ExcelColumn.class),
                                (field, annotation) -> field,
                                LinkedHashMap :: new));
    }
    private static List<? extends Formula> prepareFormulaResource(Class<?> type) {
        return ReflectionUtils.getClassAnnotationList(type, ExcelFormula.class, false).stream()
                .map(annotation -> ((ExcelFormula) annotation).formulaClass())
                .flatMap(Arrays :: stream)
                .map(ReflectionUtils :: getInstance)
                .collect(Collectors.toList());
    }
    private static Map<String, Integer> prepareFieldNameWithColumnIndexResource(Map<Field, ExcelColumn> fieldResource) {
        AtomicInteger atomicInteger = new AtomicInteger(0);
        return fieldResource.entrySet().stream().collect(
                toMap(
                        entry -> entry.getKey().getName(),
                        entry -> atomicInteger.getAndIncrement(),
                        (key, value) -> key,
                        LinkedHashMap :: new));
    }
    private static ExcelStyleResource prepareStyleResource(Class<?> type, Workbook workbook) {
        Optional<ExcelStyle> defaultStyle = ReflectionUtils.getClassAnnotationList(type, ExcelStyle.class, false)
                .stream()
                .map(annotation -> (ExcelStyle) annotation)
                .findAny();

        List<Field> fields = ReflectionUtils.getFieldWithAnnotationList(type, ExcelColumn.class, false);

        return new ExcelStyleResourceFactory(fields,defaultStyle,workbook)
                .createExcelStyleResource();
    }
    private static void validate(Map<Field, ExcelColumn> resource, Class<?> type) {
        if (resource.isEmpty()) {
            throw new NotFoundExcelColumnAnnotationException(String.format("%s has not @ExcelColumn at all", type));
        }
    }
}
