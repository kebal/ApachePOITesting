package com.company;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.ss.util.SheetUtil;

import java.io.*;
import java.util.*;



public class XLSReportGenerator {
    public static final int MERGE_COMPANY_REGION_HEIGHT = 5;
    public static final int MERGE_REPORT_REGION_HEIGHT = 4;
    public static final int TABLE_LEFT_OFFSET = 1;
    public static final int TABLE_RIGHT_OFFSET = 1;
    public static final double COMPANY_NAME_FONT_SIZE = 20;
    public static final double REPORT_NAME_FONT_SIZE = 20;
    private static final int INTIAL_CELL_WIDTH = 3000;
    private static final int FILTER_WIDTH_OFFSET = 500;

    //SPECIALLY FOR MAX
    public static final int DATA_X_OFFSET = 1;
    public static final int DATA_Y_OFFSET = MERGE_COMPANY_REGION_HEIGHT + MERGE_REPORT_REGION_HEIGHT + 1;

    private String companyName;
    private String reportName;
    private String[] columnNames;
    private List<Object[]> data;

    private HSSFWorkbook workbook;
    private HSSFSheet sheet;

    public XLSReportGenerator(String companyName, String reportName, String[] columnNames, List<Object[]> data) {
        this.companyName = companyName;
        this.reportName = reportName;
        this.columnNames = columnNames;
        this.data = data;
        workbook = new HSSFWorkbook();
        sheet = workbook.createSheet("Report");
    }

    private void generateStyle() {
        int totalColumnCount = columnNames.length + TABLE_LEFT_OFFSET + TABLE_RIGHT_OFFSET;
        for (int i = 0; i < MERGE_COMPANY_REGION_HEIGHT + MERGE_REPORT_REGION_HEIGHT + data.size() + 2; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < DATA_Y_OFFSET + data.size() + 1; j++)
                row.createCell(j);
        }
        int mergeWidth = columnNames.length;
        sheet.addMergedRegion(new CellRangeAddress(
                0,                                                      //first row (0-based)
                MERGE_COMPANY_REGION_HEIGHT - 1,                        //last row  (0-based)
                0,                                                      //first column (0-based)
                totalColumnCount - 1                                    //last column  (0-based)
        ));
        sheet.addMergedRegion(new CellRangeAddress(
                MERGE_COMPANY_REGION_HEIGHT,                            //first row (0-based)
                MERGE_COMPANY_REGION_HEIGHT,                            //last row  (0-based)
                0,                                                      //first column (0-based)
                totalColumnCount - 1                                    //last column  (0-based)
        ));
        CellRangeAddress reportNameRegion = new CellRangeAddress(
                MERGE_COMPANY_REGION_HEIGHT + 1,                          //first row (0-based)
                MERGE_COMPANY_REGION_HEIGHT + MERGE_REPORT_REGION_HEIGHT, //last row  (0-based)
                1,                                                        //first column (0-based)
                totalColumnCount - 2                                      //last column  (0-based)
        );

        CellRangeAddress leftSeparator = new CellRangeAddress(
                DATA_Y_OFFSET - MERGE_REPORT_REGION_HEIGHT,
                DATA_Y_OFFSET + data.size() + 1,
                0,
                0
        );

        CellRangeAddress rightSeparator = new CellRangeAddress(
                DATA_Y_OFFSET - MERGE_REPORT_REGION_HEIGHT,
                DATA_Y_OFFSET + data.size() + 1,
                totalColumnCount - 1,
                totalColumnCount - 1
        );

        setLeftRightBorders(reportNameRegion, CellStyle.BORDER_THIN, HSSFColor.WHITE.index);
        setPureBackGround(leftSeparator, HSSFColor.WHITE.index);
        setPureBackGround(rightSeparator, HSSFColor.WHITE.index);
        sheet.addMergedRegion(reportNameRegion);
        sheet.addMergedRegion(leftSeparator);
        sheet.addMergedRegion(rightSeparator);

        //Font for company name
        Font companyNameCellFont = workbook.createFont();
        companyNameCellFont.setColor(HSSFColor.WHITE.index);
        companyNameCellFont.setFontHeightInPoints((short) COMPANY_NAME_FONT_SIZE);
        companyNameCellFont.setFontName("Berlin Sans FB Demi");
        companyNameCellFont.setItalic(true);
        //Cell style for company name
        HSSFCellStyle companyNameCellStyle = workbook.createCellStyle();
        companyNameCellStyle.setFillForegroundColor(HSSFColor.AQUA.index);
        companyNameCellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        companyNameCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        companyNameCellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        companyNameCellStyle.setFont(companyNameCellFont);

        //Fill the first merged region with company name
        Cell cell = sheet.getRow(0).createCell(0);
        cell.setCellStyle(companyNameCellStyle);
        cell.setCellValue(companyName);

        //Font for report name
        Font reportNameCellFont = workbook.createFont();
        reportNameCellFont.setColor(HSSFColor.AQUA.index);
        reportNameCellFont.setFontHeightInPoints((short) REPORT_NAME_FONT_SIZE);
        reportNameCellFont.setFontName("Times New Roman");
        reportNameCellFont.setUnderline(HSSFFont.U_SINGLE);

        //Cell style for report name
        HSSFCellStyle reportNameCellStyle = workbook.createCellStyle();
        reportNameCellStyle.setAlignment(HSSFCellStyle.ALIGN_LEFT);
        reportNameCellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        reportNameCellStyle.setFillForegroundColor(HSSFColor.WHITE.index);
        reportNameCellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        reportNameCellStyle.setFont(reportNameCellFont);

        //Fill the first merged region with report name
        Row row = sheet.getRow(MERGE_COMPANY_REGION_HEIGHT + 1);
        row.createCell(0);
        cell = sheet.getRow(MERGE_COMPANY_REGION_HEIGHT + 1).createCell(1);
        cell.setCellStyle(reportNameCellStyle);
        cell.setCellValue(reportName);


        double totalWidth = ((COMPANY_NAME_FONT_SIZE / 1.6)/ (totalColumnCount)) * companyName.length() * 256;
        double width = totalWidth / (totalColumnCount);
        for (int i = 0; i < totalColumnCount; i++) {
            sheet.setColumnWidth(i, (int) width);
        }

        {//Fill the one line between company name and report name merge regions.
            HSSFCellStyle tempCellStyle = workbook.createCellStyle();
            tempCellStyle.setBorderTop(CellStyle.BORDER_THICK);
            tempCellStyle.setTopBorderColor(HSSFColor.WHITE.index);
            tempCellStyle.setFillForegroundColor(HSSFColor.AQUA.index);
            tempCellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
            Row tempRow = sheet.getRow(MERGE_COMPANY_REGION_HEIGHT);
            tempRow.setHeight((short) 150);
            for (int j = 0; j < totalColumnCount; j++)
                tempRow.createCell(j).setCellStyle(tempCellStyle);
        }

        //Fill data, just for testing

//        row = sheet.getRow(DATA_Y_OFFSET);
//        row.createCell(0);
//        row.createCell(1).setCellValue(columnNames[0]);
//        row.createCell(2).setCellValue(columnNames[1]);
//        row.createCell(3).setCellValue(columnNames[2]);
//        sheet.setColumnWidth(1, INTIAL_CELL_WIDTH);
//        sheet.setColumnWidth(2, INTIAL_CELL_WIDTH);
//        sheet.setColumnWidth(3,INTIAL_CELL_WIDTH);


        //Check if autoSizeColumn functions reduced the width of row (thus company name doesn't fit the cell)
        //and expand the very right and the very left cells
        int sum = 0;
        for (int i = 0; i < totalColumnCount; i++)
            sum += sheet.getColumnWidth(i);
        if (sum < totalWidth) {
            double difference = totalWidth - sum;
            sheet.setColumnWidth(0, (int) (width + difference / 2));
            sheet.setColumnWidth(totalColumnCount - 1, (int) (width + difference / 2));
        } else {
            double difference = sum - totalWidth;
            sheet.setColumnWidth(0, Math.max(2000, (int) (width - difference)));
            sheet.setColumnWidth(totalColumnCount - 1, Math.max(2000, (int) (width - difference)));
        }

        fillData();
    }

    private void setLeftRightBorders(CellRangeAddress region, short border, short color) {
        RegionUtil.setBorderLeft(border, region, sheet, workbook);
        RegionUtil.setBorderRight(border, region, sheet, workbook);
        RegionUtil.setLeftBorderColor(color, region, sheet, workbook);
        RegionUtil.setRightBorderColor(color, region, sheet, workbook);
    }

    private void setPureBackGround(CellRangeAddress region, short color) {
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillForegroundColor(color);
        cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        Cell cell = sheet.getRow(region.getFirstRow()).getCell(region.getFirstColumn());
        cell.setCellStyle(cellStyle);
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
// generateWorkbook(); - cool company title
// generateStyle(); - generate style for every cell DO NOT FORGET BORDERS
// generateReportName(); - set report name - param report type
// fillColumns(); - set data for every cell
// writeWorkBook() - save xls
*/

    private void fillData(){
        Row titleRow = sheet.getRow(DATA_Y_OFFSET);
        for(int i = DATA_X_OFFSET; i < columnNames.length + DATA_X_OFFSET;i++) {
            titleRow.createCell(i).setCellValue(columnNames[i-DATA_X_OFFSET]);
            sheet.setColumnWidth(i, INTIAL_CELL_WIDTH + FILTER_WIDTH_OFFSET);
            sheet.setAutoFilter(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, i));
        }

        int rowOffset =  DATA_Y_OFFSET + 1;
        for(int i = 0; i < data.size();i++){
            Row dataRow = sheet.getRow(rowOffset);
            dataRow.createCell(i).setCellValue("aaa");
            rowOffset++;
        }





        sheet.setAutoFilter(new CellRangeAddress(titleRow.getRowNum(),titleRow.getRowNum(),DATA_X_OFFSET,columnNames.length ));



    }


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

        XLSReportGenerator main = new XLSReportGenerator("VERY COOL PROVIDER", "CI Report",
                             new String[] { "Emp", "Emp", "Emp"}, data);
        main.createXlsFile();

    }
}
