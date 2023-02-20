package org.example.exception;

public class NotFoundExcelColumnAnnotationException extends ExcelException {
    public NotFoundExcelColumnAnnotationException(String message) {
        super(message, null);
    }
}
