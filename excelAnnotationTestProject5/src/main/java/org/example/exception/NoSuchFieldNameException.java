package org.example.exception;

public class NoSuchFieldNameException extends ExcelException{

    public NoSuchFieldNameException(String message) {
        super(message,null);
    }
}
