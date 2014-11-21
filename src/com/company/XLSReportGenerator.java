package com.company;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.RegionUtil;

import java.io.*;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.*;


public class XLSReportGenerator  {
    private static final int MERGE_COMPANY_ROW_HEIGHT = 1250;
    private static final int MERGE_REPORT_ROW_HEIGHT = 800;
    private static final int TABLE_LEFT_OFFSET = 1;
    private static final int TABLE_RIGHT_OFFSET = 1;
    private static final double COMPANY_NAME_FONT_SIZE = 24;
    private static final double REPORT_NAME_FONT_SIZE = 20;
    private static final int INITIAL_CELL_WIDTH = 3000;
    private static final int INITIAL_LETTER_WIDTH = 250;
    private static final int FILTER_WIDTH_OFFSET = 1250;
    private static final String PROVIDER_NAME = "Cauliflower";

    /**
     * Value - {@value}, column position to start filling data with .
     */
    private static final int DATA_X_OFFSET = 1;
    /**
     * Value - {@value}, row position to start filling data with .
     */
    private static final int DATA_Y_OFFSET = 3;

    private String reportName;
    private ResultSet resultSet;
    private ResultSetMetaData metaData;
    private HSSFWorkbook workbook;
    private HSSFSheet sheet;

    /**
     * Constructor.
     *

    * @param reportName  name of the report
   //  * @param columnNames array of columns names
    // * @param data        list of objects array to fill columns
     */
    public XLSReportGenerator(String reportName,ResultSet resultSet) throws SQLException{
        this.resultSet = resultSet;
        metaData = resultSet.getMetaData();
        this.reportName = reportName;
        workbook = new HSSFWorkbook();
        sheet = workbook.createSheet("Report");
        generateDocument();
    }

    /**
     * Generates look up of a document and fills it with data.
     */
    private void generateDocument() throws SQLException{

        resultSet.getRow();
        int totalColumnCount = metaData.getColumnCount() + TABLE_LEFT_OFFSET + TABLE_RIGHT_OFFSET;
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

        RegionUtil.setBorderBottom(CellStyle.BORDER_THIN, bottomSeparator, sheet, workbook);
        RegionUtil.setBottomBorderColor(HSSFColor.AQUA.index, bottomSeparator, sheet, workbook);

        RegionUtil.setBorderBottom(CellStyle.BORDER_THIN, leftSeparator, sheet, workbook);
        RegionUtil.setBottomBorderColor(HSSFColor.AQUA.index, leftSeparator, sheet, workbook);

        RegionUtil.setBorderBottom(CellStyle.BORDER_THIN, rightSeparator, sheet, workbook);
        RegionUtil.setBottomBorderColor(HSSFColor.AQUA.index, rightSeparator, sheet, workbook);

        sheet.addMergedRegion(reportNameRegion);
        sheet.addMergedRegion(leftSeparator);
        sheet.addMergedRegion(rightSeparator);
        sheet.addMergedRegion(bottomSeparator);

        //Font for company name
        Font companyNameCellFont = workbook.createFont();
        companyNameCellFont.setColor(HSSFColor.WHITE.index);
        companyNameCellFont.setFontHeightInPoints((short) COMPANY_NAME_FONT_SIZE);
        companyNameCellFont.setFontName("Segoe UI Semilight");

        //Fill the merged region with company name
        Row row = sheet.getRow(0);
        row.setHeight((short) MERGE_COMPANY_ROW_HEIGHT);
        Cell cell = row.getCell(0);
        CellUtil.setCellStyleProperty(cell, workbook, "fillForegroundColor", HSSFColor.AQUA.index);
        CellUtil.setCellStyleProperty(cell, workbook, "fillPattern", CellStyle.SOLID_FOREGROUND);
        CellUtil.setCellStyleProperty(cell, workbook, "alignment", HSSFCellStyle.ALIGN_CENTER);
        CellUtil.setCellStyleProperty(cell, workbook, "verticalAlignment", HSSFCellStyle.VERTICAL_CENTER);
        cell.getCellStyle().setFont(companyNameCellFont);
        cell.setCellValue(companyName);

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

        //Font for report name
        Font reportNameCellFont = workbook.createFont();
        reportNameCellFont.setColor(HSSFColor.AQUA.index);
        reportNameCellFont.setFontHeightInPoints((short) REPORT_NAME_FONT_SIZE);
        reportNameCellFont.setFontName("Segoe UI Semilight");

        //Fill the merged region with report name
        row = sheet.getRow(2);
        row.setHeight((short) MERGE_REPORT_ROW_HEIGHT);
        cell = row.getCell(TABLE_LEFT_OFFSET);
        CellUtil.setCellStyleProperty(cell, workbook, "fillForegroundColor", HSSFColor.WHITE.index);
        CellUtil.setCellStyleProperty(cell, workbook, "fillPattern", CellStyle.SOLID_FOREGROUND);
        CellUtil.setCellStyleProperty(cell, workbook, "alignment", HSSFCellStyle.ALIGN_LEFT);
        CellUtil.setCellStyleProperty(cell, workbook, "verticalAlignment", HSSFCellStyle.VERTICAL_BOTTOM);
        cell.getCellStyle().setFont(reportNameCellFont);
        cell.setCellValue(reportName);

        double totalWidth = ((COMPANY_NAME_FONT_SIZE / 1.6) / (totalColumnCount)) * companyName.length() * 256;
        double width = totalWidth / (totalColumnCount);
        for (int i = 0; i < totalColumnCount; i++) {
            sheet.setColumnWidth(i, (int) width);
        }

        //Set data style
        CellStyle dataStyle = workbook.createCellStyle();
        dataStyle.setAlignment(CellStyle.ALIGN_LEFT);
        dataStyle.setFillForegroundColor(HSSFColor.WHITE.index);
        dataStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        dataStyle.setBorderTop(CellStyle.BORDER_THIN);
        dataStyle.setBorderBottom(CellStyle.BORDER_THIN);
        dataStyle.setBorderLeft(CellStyle.BORDER_THIN);
        dataStyle.setBorderRight(CellStyle.BORDER_THIN);
        dataStyle.setTopBorderColor(HSSFColor.GREY_25_PERCENT.index);
        dataStyle.setBottomBorderColor(HSSFColor.GREY_25_PERCENT.index);
        dataStyle.setLeftBorderColor(HSSFColor.GREY_25_PERCENT.index);
        dataStyle.setRightBorderColor(HSSFColor.GREY_25_PERCENT.index);
        //Set column names cell style
        CellStyle columnNameCellStyle = workbook.createCellStyle();
        columnNameCellStyle.setFillForegroundColor(HSSFColor.AQUA.index);
        columnNameCellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        //Fill table with these styles
        setDataStyle(dataStyle, columnNameCellStyle);
        fillData();

        //Check if fillData() method changed values so that company name doesn't fit it's cell width.
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
        Cell cell = sheet.getRow(region.getFirstRow()).getCell(region.getFirstColumn());
        CellUtil.setCellStyleProperty(cell, workbook, "fillForegroundColor", color);
        CellUtil.setCellStyleProperty(cell, workbook, "fillPattern", CellStyle.SOLID_FOREGROUND);
    }

    private void setDataStyle(CellStyle dataStyle, CellStyle columnNameCellStyle) {
        int rowCount = DATA_Y_OFFSET + 1;
        int columnsCount = DATA_X_OFFSET;
        for (int i = 0; i < columnNames.length; i++)
            sheet.getRow(rowCount - 1).getCell(columnsCount + i).setCellStyle(columnNameCellStyle);
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.getRow(rowCount + i);
            for (int j = 0; j < columnNames.length; j++) {
                row.getCell(columnsCount + j).setCellStyle(dataStyle);
            }
        }
    }

    private double countWidthCoef(int index) {
        double res = INITIAL_CELL_WIDTH / INITIAL_LETTER_WIDTH;
        if(columnNames[index - DATA_X_OFFSET]==null)
            return 1;
        double coef = (double) (INITIAL_CELL_WIDTH / INITIAL_LETTER_WIDTH) / columnNames[index - DATA_X_OFFSET].length();
        if (coef > 4)
            res = (res / 3.9);
        else if (coef > 2)
            res = (res / 1.9);
        else if (coef < 1)
            res = columnNames[index - DATA_X_OFFSET].length();

        return res;

    }

    private void fillData() {
        //Array to save column width, to fit all inserted data.
        double[] widths = new double[columnNames.length];
        Row titleRow = sheet.getRow(DATA_Y_OFFSET);
        for (int i = DATA_X_OFFSET; i < columnNames.length + DATA_X_OFFSET; i++) {
            titleRow.getCell(i).setCellValue(columnNames[i - DATA_X_OFFSET]);
            widths[i - DATA_X_OFFSET] = countWidthCoef(i);
            //Math.max(columnNames[i - DATA_X_OFFSET].length(), INITIAL_CELL_WIDTH / INITIAL_LETTER_WIDTH);
        }
        sheet.setAutoFilter(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), DATA_X_OFFSET, columnNames.length));

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
                if(object != null)
                widths[j] = Math.max(widths[j], (double) object.toString().length());
                j++;
            }
        }

        for (int i = 1; i <= widths.length; i++) {
            sheet.setColumnWidth(i, (int) (widths[i - 1] * INITIAL_LETTER_WIDTH + FILTER_WIDTH_OFFSET));
        }
    }

    public void createXlsFile(String fileName) {
        try {
            File file = new File(fileName);
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
        data.add(new Object[]{1d, "Johnsdgdsgdsgsdg", "150d"});
        data.add(new Object[]{null, "Samsdgsdgdsg", "800000d"});
        data.add(new Object[]{null, "Deansdgsdgsdg", "700000d"});
        data.add(new Object[]{null,null ,"22222ddsgdsgdsgdsgsdgsdgsdg"});

      //  Statement stmt = con.createStatement(
//                ResultSet.TYPE_SCROLL_INSENSITIVE,
//                ResultSet.CONCUR_UPDATABLE);
//        ResultSet rs = stmt.executeQuery("SELECT a, b FROM TABLE2");



        XLSReportGenerator main = new XLSReportGenerator("VERY COOL PROVIDER", "SI Report",
                new String[]{null, "Emphhhhhhhhhhhhhhhhhhhhhhhhhhh", "Emp","ONE MORE"}, data);
        main.createXlsFile("report.xls");

    }
}
