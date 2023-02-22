package org.excel.support;

import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.Test;

import java.io.*;

public class TestSupport {

    @Test
    void test(){
        String target = "A1";
        String ROW_ABS_PATTERN = "([A-Z]+)([0-9]+)";
        String str = target.replaceAll(ROW_ABS_PATTERN, "$1\\$$2");
        System.out.println(str);
    }

    public static void workBookOutput(Workbook workbook) {
        StringBuilder stringBuilder = new StringBuilder(System.getProperty("user.dir")).append("\\src\\test\\resources");
        stringBuilder.append(File.separator);
        try (OutputStream fileOut = new FileOutputStream(stringBuilder.append("workbook.xlsx").toString())) {
            workbook.write(fileOut);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

}
