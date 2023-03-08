package com.rlung.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;
import java.util.Iterator;

public class ExcelUtils {
    public static void writeToExcel(Connection conn, String sql, String sql2, String sheetName1, String sheetName2, String outputPath) throws SQLException, IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            createSheet(conn, workbook, sql, sheetName1);
            createSheet(conn, workbook, sql2, sheetName2);
            try (FileOutputStream out = new FileOutputStream(outputPath)) {
                workbook.write(out);
            }
        }
    }
    public static void createSheet(Connection conn, XSSFWorkbook workbook, String sql, String sheetName) throws SQLException {
        int rowCount = 0;
        XSSFSheet sheet = workbook.createSheet(sheetName);

        sheet.setColumnWidth(0, (10 * 1280));
        sheet.setColumnWidth(1, (10 * 560));
        sheet.setColumnWidth(2, (10 * 1580));

        //樣式設定
        CellStyle firstRowStyle = getFirstRowStyle(workbook);
        CellStyle style1 = getStyle1(workbook);
        CellStyle style2 = getStyle2(workbook);
        CellStyle tableNamestyle1 = getColumnStyle1(workbook);
        CellStyle tableNamestyle2 = getColumnStyle2(workbook);
        CellStyle titleStyle1 = getFirstRowStyle(workbook);
        CellStyle titleStyle2 = getFirstRowStyle2(workbook);

        String currentTableName = "";
        boolean isSame = true;
        boolean isTitle = true;
        CellStyle style = style2;
        CellStyle tableNameStyle = tableNamestyle2;
        CellStyle titleStyle = titleStyle1;
        Row row2;

        PreparedStatement stmt = conn.prepareStatement(sql);
        ResultSet resultSet = stmt.executeQuery();

        try {
            while (resultSet.next()) {
                row2 = null;
                String tableN = resultSet.getString("TABLE_NAME");
                String columN = resultSet.getString("COLUMN_NAME");
                if (!tableN.equals(currentTableName)) {
                    Row row1 = sheet.createRow(rowCount++);
                    row2 = sheet.createRow(rowCount++);
                    row1.setHeightInPoints(22f);
                    row2.setHeightInPoints(30F);
                    CellRangeAddress region = new CellRangeAddress(rowCount - 1, rowCount - 1, 0, 2);
                    sheet.addMergedRegion(region);
                    isTitle = true;
                    currentTableName = tableN;
                    isSame = false;
                } else {
                    isSame = true;
                    isTitle = false;
                }
                Row row = sheet.createRow(rowCount++);
                row.setHeightInPoints(22F);
                // 在第一欄新增一個儲存格來顯示表格名稱
                if (isTitle) {
                    Cell titleCell0 = row2.createCell(0);
                    Cell titleCell1 = row2.createCell(1);
                    Cell titleCell2 = row2.createCell(2);
                    if (titleStyle == titleStyle1) {
                        titleStyle = titleStyle2;
                    } else if (titleStyle == titleStyle2) {
                        titleStyle = titleStyle1;
                    }
                    titleCell0.setCellStyle(titleStyle);
                    titleCell1.setCellStyle(titleStyle);
                    titleCell2.setCellStyle(titleStyle);
                    titleCell0.setCellValue(tableN);
                }
                if (!isSame) {
                    if (style == style1) {
                        style = style2;
                        tableNameStyle = tableNamestyle2;
                    } else if (style == style2) {
                        style = style1;
                        tableNameStyle = tableNamestyle1;
                    }
                }

                int columnCount = resultSet.getMetaData().getColumnCount();
                for (int i = 2; i <= columnCount + 1; i++) {
                    Cell cell = row.createCell(i - 2);

                    if (i > 1 && i < 4) {
                        String value = resultSet.getString(i);
                        cell.setCellValue(value);
                        if (value.equals(columN)) {
                            cell.setCellStyle(tableNameStyle);
                            continue;
                        } else {
                            cell.setCellStyle(style);
                        }
                    }
                    cell.setCellStyle(style);
                }
            }
        } catch (Exception e) {
            System.out.println(e);
            e.printStackTrace();
        }


    }


    public static String getTableNameByReadExcel(String inputPath) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(inputPath);
        Workbook inputBook = WorkbookFactory.create(fileInputStream);
        Sheet sheet = inputBook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();
        StringBuilder output = new StringBuilder();
        while (rowIterator.hasNext()) {
            Row currentRow = rowIterator.next();
            if (rowIterator.hasNext()) {
                Row nextRow = rowIterator.next();
                Cell cell = currentRow.getCell(0);
                Cell nextcell = nextRow.getCell(0);
                if (!cell.getStringCellValue().equals("") && !nextcell.getStringCellValue().equals("") && !cell.getStringCellValue().equals("Table name")) {
                    output.append("'").append(cell.getStringCellValue()).append("',");
                } else if (!cell.getStringCellValue().equals("") && nextcell.getStringCellValue().equals("") && !cell.getStringCellValue().equals("Table name")) {
                    output.append("'").append(cell.getStringCellValue()).append("'");
                }
            }
        }
        return output.toString();
    }

    public static XSSFCellStyle getFirstRowStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        //邊框線
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());

        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 14);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        XSSFColor color = new XSSFColor(new java.awt.Color(74, 134, 232));
        style.setFillForegroundColor(color);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);


        return style;
    }

    public static XSSFCellStyle getFirstRowStyle2(XSSFWorkbook workbook) {
        XSSFCellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        //邊框線
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());

        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 14);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        XSSFColor color = new XSSFColor(new java.awt.Color(255, 153, 0));
        style.setFillForegroundColor(color);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);


        return style;
    }


    public static XSSFCellStyle getStyle1(XSSFWorkbook workbook) {
        XSSFCellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        //邊框線
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());

        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 14);
        font.setColor(IndexedColors.BLACK.getIndex());
        style.setFont(font);
        XSSFColor color = new XSSFColor(new java.awt.Color(255, 242, 204));
        style.setFillForegroundColor(color);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        style.setVerticalAlignment(VerticalAlignment.CENTER);


        return style;
    }

    public static XSSFCellStyle getColumnStyle1(XSSFWorkbook workbook) {
        XSSFCellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        //邊框線
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());

        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 14);
        font.setColor(IndexedColors.BLACK.getIndex());
        style.setFont(font);
        XSSFColor color = new XSSFColor(new java.awt.Color(249, 203, 156));
        style.setFillForegroundColor(color);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        style.setVerticalAlignment(VerticalAlignment.CENTER);


        return style;
    }

    public static XSSFCellStyle getStyle2(XSSFWorkbook workbook) {
        XSSFCellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        //邊框線
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());

        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 14);
        font.setColor(IndexedColors.BLACK.getIndex());
        style.setFont(font);
        XSSFColor color = new XSSFColor(new java.awt.Color(207, 226, 243));
        style.setFillForegroundColor(color);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        style.setVerticalAlignment(VerticalAlignment.CENTER);


        return style;
    }

    public static XSSFCellStyle getColumnStyle2(XSSFWorkbook workbook) {
        XSSFCellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        //邊框線
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());

        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 14);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        XSSFColor color = new XSSFColor(new java.awt.Color(109, 158, 235));
        style.setFillForegroundColor(color);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        style.setVerticalAlignment(VerticalAlignment.CENTER);


        return style;
    }

    public static XSSFCellStyle getBlankStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        //邊框線
        style.setBorderBottom(BorderStyle.THICK);
        style.setBorderTop(BorderStyle.THICK);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());

        style.setVerticalAlignment(VerticalAlignment.CENTER);

        return style;
    }
}
