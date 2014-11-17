package com.company;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.awt.*;
import java.io.*;
import java.util.*;

public class Main {

    public static void generateXLS(String[] columnNames, Map<Integer, Object[]> data) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Report");

        Row namesRow = sheet.createRow(0);
        HSSFCellStyle mainStyle = workbook.createCellStyle();
        mainStyle.setFillForegroundColor(HSSFColor.TEAL.index);
        mainStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        mainStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        for (int i = 0; i < columnNames.length; i++) {
            Cell cell = namesRow.createCell(i);
            cell.setCellStyle(mainStyle);
            cell.setCellValue(columnNames[i]);
        }

        int rowNum = 0;
        for (Integer key : data.keySet()) {
            Row row = sheet.createRow(rowNum);
            Object [] objArr = data.get(key);
            int cellNum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellNum++);
                if(obj instanceof Date)
                    cell.setCellValue((Date)obj);
                else if(obj instanceof Boolean)
                    cell.setCellValue((Boolean)obj);
                else if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Double)
                    cell.setCellValue((Double)obj);
            }
            sheet.autoSizeColumn(rowNum);
            rowNum++;
        }

    }

    public static void writeWorkBook(HSSFWorkbook workbook) {
        try {
            FileOutputStream out = new FileOutputStream(new File("report.xls"));
            workbook.write(out);
            out.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        Map<Integer, Object[]> data = new HashMap<Integer, Object[]>();
        data.put(1, new Object[] {1d, "John", 15000000000d});
        data.put(2, new Object[] {2d, "Sam", 800000d});
        data.put(3, new Object[] {3d, "Dean", 700000d});

        generateXLS(new String[] { "Emp No.", "Name", "Salary"}, data);

    }
}
