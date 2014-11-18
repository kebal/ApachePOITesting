package com.company;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import java.io.*;
import java.util.*;


public class XLSReportGenerator {


    public static final int MERGE_COMPANY_ROW_HEIGHT = 1250;
    public static final int MERGE_REPORT_ROW_HEIGHT = 750;
    public static final int TABLE_LEFT_OFFSET = 1;
    public static final int TABLE_RIGHT_OFFSET = 1;
    public static final double COMPANY_NAME_FONT_SIZE = 20;
    public static final double REPORT_NAME_FONT_SIZE = 20;
    private static final int INITIAL_CELL_WIDTH = 3000;
    private static final int INITIAL_LETTER_WIDTH = 250;
    private static final int FILTER_WIDTH_OFFSET = 1250;

    //SPECIALLY FOR MAX

    /** Value - {@value}, column position to start filling data with .*/
    public static final int DATA_X_OFFSET = 1;
    /** Value - {@value}, row position to start filling data with .*/
    public static final int DATA_Y_OFFSET = 3;

    private String companyName;
    private String reportName;
    private String[] columnNames;
    private List<Object[]> data;

    private HSSFWorkbook workbook;
    private HSSFSheet sheet;

    /**
     * Constructor.
     *
     * @param companyName name of the company created report
     * @param reportName  name of the report
     * @param columnNames array of columns names
     * @param data        list of objects array to fill columns
     */
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
        for (int i = 0; i < DATA_Y_OFFSET + data.size() + 2; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < totalColumnCount; j++)
                row.createCell(j);
        }
        int mergeWidth = columnNames.length;
        sheet.addMergedRegion(new CellRangeAddress(
                0,                                                      //first row (0-based)
                0,                        //last row  (0-based)
                0,                                                      //first column (0-based)
                totalColumnCount - 1                                    //last column  (0-based)
        ));
        sheet.addMergedRegion(new CellRangeAddress(
                1,                            //first row (0-based)
                1,                            //last row  (0-based)
                0,                                                      //first column (0-based)
                totalColumnCount - 1                                    //last column  (0-based)
        ));
        CellRangeAddress reportNameRegion = new CellRangeAddress(
                2,                          //first row (0-based)
                2, //last row  (0-based)
                1,                                                        //first column (0-based)
                totalColumnCount - 2                                      //last column  (0-based)
        );

        CellRangeAddress leftSeparator = new CellRangeAddress(
                DATA_Y_OFFSET - 1,
                DATA_Y_OFFSET + data.size() + 1,
                0,
                0
        );

        CellRangeAddress rightSeparator = new CellRangeAddress(
                DATA_Y_OFFSET - 1,
                DATA_Y_OFFSET + data.size() + 1,
                totalColumnCount - 1,
                totalColumnCount - 1
        );

        CellRangeAddress bottomSeparator = new CellRangeAddress(
                DATA_Y_OFFSET + data.size() + 1,
                DATA_Y_OFFSET + data.size() + 1,
                TABLE_LEFT_OFFSET,
                totalColumnCount - TABLE_RIGHT_OFFSET - 1
        );

        setLeftRightBorders(reportNameRegion, CellStyle.BORDER_THIN, HSSFColor.WHITE.index);
        setPureBackGround(rightSeparator, HSSFColor.WHITE.index);
        setPureBackGround(leftSeparator, HSSFColor.WHITE.index);
        setPureBackGround(bottomSeparator, HSSFColor.WHITE.index);
        sheet.addMergedRegion(reportNameRegion);
        sheet.addMergedRegion(leftSeparator);
        sheet.addMergedRegion(rightSeparator);
        sheet.addMergedRegion(bottomSeparator);

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
        Row row = sheet.getRow(0);
        row.setHeight((short) MERGE_COMPANY_ROW_HEIGHT);
        Cell cell = row.getCell(0);
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
        row = sheet.getRow(2);
        row.setHeight((short)MERGE_REPORT_ROW_HEIGHT);
        cell = row.getCell(TABLE_LEFT_OFFSET);
        cell.setCellStyle(reportNameCellStyle);
        cell.setCellValue(reportName);


        double totalWidth = ((COMPANY_NAME_FONT_SIZE / 1.6) / (totalColumnCount)) * companyName.length() * 256;
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
            Row tempRow = sheet.getRow(1);
            tempRow.setHeight((short) 150);
            for (int j = 0; j < totalColumnCount; j++)
                tempRow.getCell(j).setCellStyle(tempCellStyle);
        }

        fillData();

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

    private void fillData() {
        int [] widths = new int[columnNames.length];
        Row titleRow = sheet.getRow(DATA_Y_OFFSET);
        for (int i = DATA_X_OFFSET; i < columnNames.length + DATA_X_OFFSET; i++) {
            titleRow.getCell(i).setCellValue(columnNames[i - DATA_X_OFFSET]);
            widths[i - DATA_X_OFFSET] = Math.max(columnNames[i - DATA_X_OFFSET].length(),INITIAL_CELL_WIDTH / INITIAL_LETTER_WIDTH);
        }
        sheet.setAutoFilter(new CellRangeAddress(titleRow.getRowNum(),titleRow.getRowNum(),DATA_X_OFFSET,columnNames.length));

        int rowCount = DATA_Y_OFFSET + 1;
        int columnsCount = DATA_X_OFFSET;

        for (int i = 0; i < data.size(); i++) {
            Object[] dataInfo = data.get(i);
            Row dataRow = sheet.getRow(rowCount + i);
            int j = 0;
            for (Object object : dataInfo) {
                Cell cell = dataRow.getCell(columnsCount + j);
                if (object instanceof Double) {
                    cell.setCellValue((Double) object);
                } else if (object instanceof String) {
                    cell.setCellValue((String) object);
                } else if (object instanceof Date) {
                    cell.setCellValue((Date) object);
                } else if (object instanceof Boolean) {
                    cell.setCellValue((Date) object);
                }
                widths[j] = Math.max(widths[j], object.toString().length());
                j++;
            }
        }

        for (int i = 1; i <= widths.length; i++) {
            sheet.setColumnWidth(i, widths[i - 1] * INITIAL_LETTER_WIDTH + FILTER_WIDTH_OFFSET);
        }
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
        data.add(new Object[]{1d, "John", "150d"});
        data.add(new Object[]{2d, "Sam", "800000d"});
        data.add(new Object[]{3d, "Dean", "700000d"});
        data.add(new Object[]{5d, "Max", "22222d"});

        XLSReportGenerator main = new XLSReportGenerator("VERY VERY VERY COOL PROVIDER", "SI Report",
                new String[]{"Emp", "Emp", "Emp"}, data);
        main.createXlsFile();

    }
}
