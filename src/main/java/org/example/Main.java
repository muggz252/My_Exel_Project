package org.example;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Main {
    public static void main(String[] args) {
        List<Client> list = new ArrayList<>();
        list.add(new Client("Max","Saint-Petersburg"));
        list.add(new Client("Petr","Moscow"));
        list.add(new Client("Inna","Ekaterinburg"));

        Map<Integer, String> map = new HashMap<>();
        Workbook wb = new HSSFWorkbook();
        listWriter(list, wb);
        System.out.println(Parser.parse("my1.xls"));
    }
    public static void listWriter(List<Client> list,Workbook wb) {
        Sheet sheet = wb.createSheet("clients");
        for (int i = 0; i < list.size(); i++) {
            Row row = sheet.createRow(i);
            Cell cell = row.createCell(0);
            cell.setCellValue(list.get(i).name);
            Cell cell1 = row.createCell(1);
            cell1.setCellValue(list.get(i).town);
            try (FileOutputStream fos = new FileOutputStream("my1.xls")) {
                wb.write(fos);
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }
    }
}