package com.company;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.*;
import java.util.*;



public class Main {
    public static final int MERGE_REGION_HEIGHT = 5;

    private String companyName;
    private String reportName;
    private String[] columnNames;
    private List<Object[]> data;

    private HSSFWorkbook workbook;
    private HSSFSheet sheet;

    public Main(String companyName, String reportName, String[] columnNames, List<Object[]> data) {
        this.companyName = companyName;
        this.reportName = reportName;
        this.columnNames = columnNames;
        this.data = data;
        workbook = new HSSFWorkbook();
        sheet = workbook.createSheet("Report");
    }

    private void generateStyle() {
        int mergeWidth = columnNames.length;
        sheet.addMergedRegion(new CellRangeAddress(
                0,                       //first row (0-based)
                MERGE_REGION_HEIGHT - 1, //last row  (0-based)
                0,                       //first column (0-based)
                mergeWidth               //last column  (0-based)
        ));
        sheet.addMergedRegion(new CellRangeAddress(
                MERGE_REGION_HEIGHT,          //first row (0-based)
                MERGE_REGION_HEIGHT * 2 - 1,          //last row  (0-based)
                0,          //first column (0-based)
                mergeWidth  //last column  (0-based)
        ));

        HSSFCellStyle companyNameCellStyle = workbook.createCellStyle();
        companyNameCellStyle.setFillForegroundColor(HSSFColor.AQUA.index);
        companyNameCellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        Row row = sheet.createRow(0);
    }
/*
Lexa:
generateWorkbook();
generateStyle();
Max:
* rename Report
*
* maybe constructor
*
* Report( string name,+ few params )
* generateReportName(); - set report name - param report type
  fillColumns(); - set data for every cell
  writeWorkBook() - save xls
*
* */

/*
    public void generateXLS(String[] columnNames, List<Object[]> data) {

    // generateWorkbook(); - cool company title
    // generateStyle(); - generate style for every cell DO NOT FORGET BORDERS
    // generateReportName(); - set report name - param report type
    // fillColumns(); - set data for every cell
    // writeWorkBook() - save xls


        //Main row with column titles
        Row namesRow = sheet.createRow(0);
        //Style for cells in this main row
        HSSFCellStyle mainStyle = workbook.createCellStyle();
        //Font for text in this row
        Font font = workbook.createFont();
        font.setColor(HSSFColor.WHITE.index);
        font.setFontHeightInPoints((short) 15);
        mainStyle.setFont(font);
        //Fill
        mainStyle.setFillForegroundColor(HSSFColor.AQUA.index);
        mainStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        //Alignment
        mainStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        mainStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        //Borders
        mainStyle.setBorderBottom(CellStyle.BORDER_THICK);
        mainStyle.setBorderLeft(CellStyle.BORDER_THICK);
        mainStyle.setBorderRight(CellStyle.BORDER_THICK);
        //Fill main row with data from columnNames

        for (int i = 0; i < 7; i++) {
            Cell cell = namesRow.createCell(i);
            if(i == 6)cell.setCellStyle(mainStyle);
           // cell.setCellValue(columnNames[i]);
        }
        sheet.autoSizeColumn(COLUMN);
        
        //Style for cells with data
//        HSSFCellStyle commonStyle = workbook.createCellStyle();
//        commonStyle.setAlignment(CellStyle.ALIGN_LEFT);
//        //Borders size
//        commonStyle.setBorderTop(CellStyle.BORDER_THIN);
//        commonStyle.setBorderBottom(CellStyle.BORDER_THIN);
//        commonStyle.setBorderLeft(CellStyle.BORDER_THIN);
//        commonStyle.setBorderRight(CellStyle.BORDER_THIN);
//        //Borders color
//        commonStyle.setTopBorderColor(HSSFColor.GREY_50_PERCENT.index);
//        commonStyle.setBottomBorderColor(HSSFColor.GREY_50_PERCENT.index);
//        commonStyle.setLeftBorderColor(HSSFColor.GREY_50_PERCENT.index);
//        commonStyle.setRightBorderColor(HSSFColor.GREY_50_PERCENT.index);
//        //Fill columns with data
//        int rowNum = 1;
//        for (Integer key : data.keySet()) {
//            Row row = sheet.createRow(rowNum);
//            Object [] objArr = data.get(key);
//            int cellNum = 0;
//            for (Object obj : objArr) {
//                Cell cell = row.createCell(cellNum++);
//                cell.setCellStyle(commonStyle);
//                if(obj instanceof Date)
//                    cell.setCellValue((Date)obj);
//                else if(obj instanceof Boolean)
//                    cell.setCellValue((Boolean)obj);
//                else if(obj instanceof String)
//                    cell.setCellValue((String)obj);
//                else if(obj instanceof Double)
//                    cell.setCellValue((Double)obj);
//            }
//            sheet.autoSizeColumn(rowNum++);
//        }
        writeWorkBook(workbook);
}
*/

    public void createXlsFile() {
        generateStyle();
        try {
            File file = new File("report.xls");
            FileOutputStream out = new FileOutputStream(file);
            workbook.write(out);
            out.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        List<Object[]> data = new ArrayList<Object[]>();
        data.add(new Object[]{1d, "John", 15000000000d});
        data.add(new Object[]{2d, "Sam", 800000d});
        data.add(new Object[]{3d, "Dean", 700000d});
        data.add(new Object[]{5d, "Max", 22222d});

        Main main = new Main("VERY COOL PROVIDER COMPANY", "CI Report",
                             new String[] { "Employee number", "Employee name", "Salary"}, data);
        main.createXlsFile();

    }
}
