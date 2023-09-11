package org.example;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class Parser {
    public static String parse(String input) {
        StringBuilder sb = new StringBuilder();
        InputStream in;
        HSSFWorkbook wb;
        try {
            in = new FileInputStream(input);
            wb = new HSSFWorkbook(in);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        Sheet sheet = wb.getSheetAt(0);
        Iterator<Row> it = sheet.iterator();
        while (it.hasNext()) {
            Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            while (cells.hasNext()) {
                Cell cell = cells.next();
                int cellType = cell.getCellType();
                switch (cellType) {
                    case Cell.CELL_TYPE_STRING -> sb.append(" ").append(cell.getStringCellValue());
                    case Cell.CELL_TYPE_NUMERIC -> sb.append("{").append(cell.getNumericCellValue()).append("}");
                    default -> sb.append("|");
                }
            }
            sb.append('\n');
        }
        return sb.toString();
    }
}
