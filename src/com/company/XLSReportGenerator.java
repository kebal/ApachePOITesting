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
import java.sql.Types;
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

    private enum XMLType {
        Number,
        Text,
        Bool,
        Date,
        ToString
    }

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
        int totalColumnCount = metaData.getColumnCount() + TABLE_LEFT_OFFSET + TABLE_RIGHT_OFFSET;
        double totalWidth = ((COMPANY_NAME_FONT_SIZE / 1.6) / (totalColumnCount)) * PROVIDER_NAME.length() * 256;
        double width = totalWidth / (totalColumnCount);
        for (int i = 0; i < totalColumnCount; i++) {
            sheet.setColumnWidth(i, (int) width);
        }
        int rowsInserted = fillData();

        int mergeWidth = metaData.getColumnCount();
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
                DATA_Y_OFFSET + rowsInserted + 1,
                0,
                0
        );

        CellRangeAddress rightSeparator = new CellRangeAddress(
                DATA_Y_OFFSET - 1,
                DATA_Y_OFFSET + rowsInserted + 1,
                totalColumnCount - 1,
                totalColumnCount - 1
        );

        CellRangeAddress bottomSeparator = new CellRangeAddress(
                DATA_Y_OFFSET + rowsInserted + 1,
                DATA_Y_OFFSET + rowsInserted + 1,
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
        Row row = CellUtil.getRow(0, sheet);
        row.setHeight((short) MERGE_COMPANY_ROW_HEIGHT);
        Cell cell = CellUtil.getCell(row, 0);
        CellUtil.setCellStyleProperty(cell, workbook, "fillForegroundColor", HSSFColor.AQUA.index);
        CellUtil.setCellStyleProperty(cell, workbook, "fillPattern", CellStyle.SOLID_FOREGROUND);
        CellUtil.setCellStyleProperty(cell, workbook, "alignment", HSSFCellStyle.ALIGN_CENTER);
        CellUtil.setCellStyleProperty(cell, workbook, "verticalAlignment", HSSFCellStyle.VERTICAL_CENTER);
        cell.getCellStyle().setFont(companyNameCellFont);
        cell.setCellValue(PROVIDER_NAME);

        {//Fill the one line between company name and report name merge regions.
            HSSFCellStyle tempCellStyle = workbook.createCellStyle();
            tempCellStyle.setBorderTop(CellStyle.BORDER_THICK);
            tempCellStyle.setTopBorderColor(HSSFColor.WHITE.index);
            tempCellStyle.setFillForegroundColor(HSSFColor.AQUA.index);
            tempCellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
            Row tempRow = CellUtil.getRow(1, sheet);
            tempRow.setHeight((short) 150);
            for (int j = 0; j < totalColumnCount; j++)
                CellUtil.getCell(tempRow, j).setCellStyle(tempCellStyle);
        }

        //Font for report name
        Font reportNameCellFont = workbook.createFont();
        reportNameCellFont.setColor(HSSFColor.AQUA.index);
        reportNameCellFont.setFontHeightInPoints((short) REPORT_NAME_FONT_SIZE);
        reportNameCellFont.setFontName("Segoe UI Semilight");

        //Fill the merged region with report name
        row = CellUtil.getRow(2, sheet);
        row.setHeight((short) MERGE_REPORT_ROW_HEIGHT);
        cell = CellUtil.getCell(row, TABLE_LEFT_OFFSET);
        CellUtil.setCellStyleProperty(cell, workbook, "fillForegroundColor", HSSFColor.WHITE.index);
        CellUtil.setCellStyleProperty(cell, workbook, "fillPattern", CellStyle.SOLID_FOREGROUND);
        CellUtil.setCellStyleProperty(cell, workbook, "alignment", HSSFCellStyle.ALIGN_LEFT);
        CellUtil.setCellStyleProperty(cell, workbook, "verticalAlignment", HSSFCellStyle.VERTICAL_BOTTOM);
        cell.getCellStyle().setFont(reportNameCellFont);
        cell.setCellValue(reportName);



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
        setDataStyle(dataStyle, columnNameCellStyle, rowsInserted);

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
        Row row = CellUtil.getRow(region.getFirstRow(), sheet);
        Cell cell = CellUtil.getCell(row, region.getFirstColumn());
        CellUtil.setCellStyleProperty(cell, workbook, "fillForegroundColor", color);
        CellUtil.setCellStyleProperty(cell, workbook, "fillPattern", CellStyle.SOLID_FOREGROUND);
    }

    /**
     *
     * @param dataStyle
     * @param columnNameCellStyle
     * @param rowsInserted How many rows inserted fillData method
     * @throws SQLException
     */
    private void setDataStyle(CellStyle dataStyle, CellStyle columnNameCellStyle, int rowsInserted) throws SQLException {
        int rowCount = DATA_Y_OFFSET + 1;
        int columnsCount = DATA_X_OFFSET;
        for (int i = 0; i < metaData.getColumnCount(); i++) {
            Row row = CellUtil.getRow(rowCount - 1, sheet);
            Cell cell = CellUtil.getCell(row, columnsCount + i);
            cell.setCellStyle(columnNameCellStyle);
        }
        for (int i = 0; i < rowsInserted - 1; i++) {//minus one because we don't take into consederation column names row
            Row row = CellUtil.getRow(rowCount + i, sheet);
            for (int j = 0; j < metaData.getColumnCount(); j++) {
                CellUtil.getCell(row, columnsCount + j).setCellStyle(dataStyle);
            }
        }
    }

    private double countWidthCoef(int index) throws SQLException {
        double res = INITIAL_CELL_WIDTH / INITIAL_LETTER_WIDTH;
//        if(metaData.getColumnLabel[index - DATA_X_OFFSET]==null)
//            return 1;
        double coef = (double) (INITIAL_CELL_WIDTH / INITIAL_LETTER_WIDTH) / metaData.getColumnLabel(index + 1).length();
        if (coef > 4)
            res = (res / 3.9);
        else if (coef > 2)
            res = (res / 1.9);
        else if (coef < 1)
            res = metaData.getColumnLabel(index + 1).length();
        return res;

    }

    /**
     *
     * @return How many rows were inserted
     * @throws SQLException
     */
    private int fillData() throws SQLException {


        int rowsInserted = 1;//1 for column names row
        //Array to save column width, to fit all inserted data.
        double[] widths = new double[metaData.getColumnCount()];
        int columnsCount = DATA_X_OFFSET - 1;
        int rowCount = DATA_Y_OFFSET + 1;
        Row titleRow = CellUtil.getRow(DATA_Y_OFFSET, sheet);

        for (int i = 0; i < metaData.getColumnCount(); i++) {
            CellUtil.getCell(titleRow, i + columnsCount).setCellValue(metaData.getColumnLabel(i + 1));
            widths[i] = countWidthCoef(i);
            //Math.max(columnNames[i ].length(), INITIAL_CELL_WIDTH / INITIAL_LETTER_WIDTH);
        }
        sheet.setAutoFilter(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), DATA_X_OFFSET, metaData.getColumnCount()));

        while (resultSet.next()) {
            Row dataRow = CellUtil.getRow(rowCount, sheet);
            rowCount++;
            for (int j = 0; j < metaData.getColumnCount(); j++) {
                Cell cell = CellUtil.getCell(dataRow, columnsCount + j);
                XMLType type = getTypeID(metaData.getColumnType(j + 1));
                Object object = resultSet.getObject(j + 1);
                switch (type) {
                    case Text:
                        cell.setCellValue((String) object);
                        break;
                    case Number:
                        cell.setCellValue((Double) object);
                        break;
                    case Date:
                        cell.setCellValue((Date) object);
                        break;
                    case Bool:
                        cell.setCellValue((Boolean) object);
                        break;
                    case ToString:
                        cell.setCellValue(object.toString());
                        break;
                }
            }
            rowsInserted++;
        }
        for (int i = 1; i <= widths.length; i++) {
            sheet.setColumnWidth(i, (int) (widths[i - 1] * INITIAL_LETTER_WIDTH + FILTER_WIDTH_OFFSET));
        }
        return rowsInserted;
    }


    /**
     * @param typeCode sql type id
     * @return
     */
    private XMLType getTypeID(int typeCode) {
        switch (typeCode) {
            case Types.DOUBLE:
            case Types.FLOAT:
            case Types.INTEGER:
            case Types.REAL:
            case Types.NUMERIC:
                return XMLType.Number;
            case Types.VARCHAR:
            case Types.CHAR:
            case Types.LONGNVARCHAR:
                return XMLType.Text;
            case Types.BOOLEAN:
                return XMLType.Bool;
            case Types.DATE:
                return XMLType.Date;
            default:
                return XMLType.ToString;
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
}