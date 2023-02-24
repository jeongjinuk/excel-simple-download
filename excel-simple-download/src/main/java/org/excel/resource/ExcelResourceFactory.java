package org.excel.resource;

import org.apache.poi.ss.usermodel.Workbook;
import org.excel.ExcelColumn;
import org.excel.ExcelFormula;
import org.excel.ExcelStyle;
import org.excel.excel_file.formula.Formula;
import org.excel.exception.NotFoundExcelColumnAnnotationException;

import java.lang.reflect.Field;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Consumer;
import java.util.stream.Stream;

import static java.util.stream.Collectors.toMap;

public final class ExcelResourceFactory {
    private ExcelResourceFactory() {
    }
    public static ExcelResource prepareExcelResource(Class<?> type, Workbook workbook) {
        Map<Field, ExcelColumn> fieldResource = prepareFieldResource(type);

        validate(fieldResource, type);

        List<Formula> formulaResource = prepareFormulaResource(type);

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
    private static List<Formula> prepareFormulaResource(Class<?> type) {
        List<Formula> formulas = new ArrayList<>();
        Consumer<ExcelFormula> consumer = (excelFormula) -> Stream.of(excelFormula)
                .map(ExcelFormula::formulaClass)
                .flatMap(Arrays::stream)
                .map(ReflectionUtils::getInstance)
                .forEach(formulas::add);

        ReflectionUtils.getOptionalClassAnnotation(type, ExcelFormula.class, false)
                .ifPresent(consumer::accept);
        return formulas;
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
        Optional<ExcelStyle> defaultStyle = ReflectionUtils.getOptionalClassAnnotation(type, ExcelStyle.class, false);
        List<Field> fields = ReflectionUtils.getFieldWithAnnotationList(type, ExcelColumn.class, false);
        return new ExcelStyleResourceFactory(fields, defaultStyle, workbook)
                .createExcelStyleResource();
    }
    private static void validate(Map<Field, ExcelColumn> resource, Class<?> type) {
        if (resource.isEmpty()) {
            throw new NotFoundExcelColumnAnnotationException(String.format("%s has not @ExcelColumn at all", type));
        }
    }
}
