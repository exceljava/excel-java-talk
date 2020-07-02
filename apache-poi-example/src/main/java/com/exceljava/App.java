/**
 * Example showing how to create an Excel workbook using
 * Apache POI.
 *
 * Demonstrates writing a simple worksheet with no formatting,
 * one with cell formatting, and one including formulas.
 *
 * The values in the report are meaningless!
 *
 */
package com.exceljava;

import com.opencsv.bean.CsvToBeanBuilder;
import org.apache.commons.io.input.BOMInputStream;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.List;

public class App 
{
    private final static String[] columns = {
            "Portfolio",
            "Security",
            "Notional",
            "Market Value",
            "Delta %",
            "Delta $",
            "1Y",
            "3Y",
            "5Y",
            "10Y",
            "15Y"
    };

    public static void main( String[] args ) throws IOException {
        // Read the CSV file to simulate some report that we might be calculating
        // or reading from a database.
        List<Record> records;
        try (FileInputStream fs = new FileInputStream("data/input.csv")) {
            records = new CsvToBeanBuilder(new InputStreamReader(new BOMInputStream(fs), StandardCharsets.UTF_8))
                    .withType(Record.class)
                    .build()
                    .parse();
        }

        // Create a new workbook and a new sheet
        XSSFWorkbook workbook = new XSSFWorkbook();

        // Write the report to the workbook
        XSSFSheet basicSheet = workbook.createSheet("Basic");
        writeBasicReport(basicSheet, records);

        // Write the formatted report to the workbook
        XSSFSheet formattedSheet = workbook.createSheet("Formatted");
        writeFormattedReport(workbook, formattedSheet, records);

        // Write the formatted report to the workbook
        XSSFSheet withSumsSheet = workbook.createSheet("With Sum");
        writeFormattedReportWithSums(workbook, withSumsSheet, records);

        // Write the workbook to a file
        try (FileOutputStream out = new FileOutputStream(new File("data/output.xlsx"))) {
            workbook.write(out);
            System.out.println("Excel workbook saved.");
        }
    }

    private static int writeBasicReport(XSSFSheet sheet, List<Record> records) {
        // Set the first row to the column headers
        int currentRow = 0;
        XSSFRow headerRow = sheet.createRow(currentRow++);

        for (int i = 0; i < columns.length; i++) {
            headerRow.createCell(i).setCellValue(columns[i]);
        }

        // Write out each record to a new row in the sheet
        for (Record record: records) {
            XSSFRow row = sheet.createRow(currentRow++);

            row.createCell(0).setCellValue(record.getPortfolio());
            row.createCell(1).setCellValue(record.getSecurity());
            row.createCell(2).setCellValue(record.getNotional());
            row.createCell(3).setCellValue(record.getMarketValue());
            row.createCell(4).setCellValue(record.getDeltaPct());
            row.createCell(5).setCellValue(record.getDeltaValue());
            row.createCell(6).setCellValue(record.getRiskFactor1y());
            row.createCell(7).setCellValue(record.getRiskFactor3y());
            row.createCell(8).setCellValue(record.getRiskFactor5y());
            row.createCell(9).setCellValue(record.getRiskFactor10y());
            row.createCell(10).setCellValue(record.getRiskFactor15y());
        }

        return currentRow;
    }

    private static int writeFormattedReport(XSSFWorkbook workbook, XSSFSheet sheet, List<Record> records) {
        // Create some styles to use to format the values
        XSSFFont whiteFont = workbook.createFont();
        whiteFont.setFontName(XSSFFont.DEFAULT_FONT_NAME);
        whiteFont.setColor(IndexedColors.WHITE.getIndex());

        XSSFCellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(32, 55, 100), new DefaultIndexedColorMap()));
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setFont(whiteFont);

        XSSFCellStyle amountStyle = workbook.createCellStyle();
        DataFormat amountFormat = workbook.createDataFormat();
        amountStyle.setDataFormat(amountFormat.getFormat("[$$-en-US]#,##0_ ;-[$$-en-US]#,##0"));

        XSSFCellStyle percentageStyle = workbook.createCellStyle();
        DataFormat percentageFormat = workbook.createDataFormat();
        percentageStyle.setDataFormat(percentageFormat.getFormat("0.00%"));

        // Set the first row to the column headers
        int currentRow = 0;
        XSSFRow headerRow = sheet.createRow(currentRow++);


        for (int i = 0; i < columns.length; i++) {
            XSSFCell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerStyle);
        }

        // Write out each record to a new row in the sheet
        for (Record record: records) {
            XSSFRow row = sheet.createRow(currentRow++);

            XSSFCell portfolioCell = row.createCell(0);
            portfolioCell.setCellValue(record.getPortfolio());

            XSSFCell securityCell = row.createCell(1);
            securityCell.setCellValue(record.getSecurity());

            XSSFCell notionalCell = row.createCell(2);
            notionalCell.setCellValue(record.getNotional());
            notionalCell.setCellStyle(amountStyle);

            XSSFCell marketValueCell = row.createCell(3);
            marketValueCell.setCellValue(record.getMarketValue());
            marketValueCell.setCellStyle(amountStyle);

            XSSFCell deltaPctCell = row.createCell(4);
            deltaPctCell.setCellValue(record.getDeltaPct());
            deltaPctCell.setCellStyle(percentageStyle);

            XSSFCell deltaValueCell = row.createCell(5);
            deltaValueCell.setCellValue(record.getDeltaValue());
            deltaValueCell.setCellStyle(amountStyle);

            XSSFCell riskFactor1yCell = row.createCell(6);
            riskFactor1yCell.setCellValue(record.getRiskFactor1y());
            riskFactor1yCell.setCellStyle(amountStyle);

            XSSFCell riskFactor3yCell = row.createCell(7);
            riskFactor3yCell.setCellValue(record.getRiskFactor3y());
            riskFactor3yCell.setCellStyle(amountStyle);

            XSSFCell riskFactor5yCell = row.createCell(8);
            riskFactor5yCell.setCellValue(record.getRiskFactor5y());
            riskFactor5yCell.setCellStyle(amountStyle);

            XSSFCell riskFactor10yCell = row.createCell(9);
            riskFactor10yCell.setCellValue(record.getRiskFactor10y());
            riskFactor10yCell.setCellStyle(amountStyle);

            XSSFCell riskFactor15yCell = row.createCell(10);
            riskFactor15yCell.setCellValue(record.getRiskFactor15y());
            riskFactor15yCell.setCellStyle(amountStyle);
        }

        // Auto-size all of the columns
        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        return currentRow;
    }

    private static int writeFormattedReportWithSums(XSSFWorkbook workbook, XSSFSheet sheet, List<Record> records) {
        // Write the formatted report with no sum line
        int currentRow = writeFormattedReport(workbook, sheet, records);
        XSSFRow lastRow = sheet.getRow(currentRow - 1);

        XSSFFont boldFont = workbook.createFont();
        boldFont.setFontName(XSSFFont.DEFAULT_FONT_NAME);
        boldFont.setBold(true);

        // Add the sums for the columns, skipping the first 5
        XSSFRow row = sheet.createRow(currentRow++);
        for (int i = 5; i < columns.length; i++) {
            XSSFCell cell = row.createCell(i);
            String colName = CellReference.convertNumToColString(i);
            String reference = colName + 2 + ":" + colName + (lastRow.getRowNum() + 1);
            cell.setCellFormula("SUM(" + reference + ")");

            // Set the style to be the same as the row above but in bold
            XSSFCell cellAbove = lastRow.getCell(i);
            XSSFCellStyle baseStyle = cellAbove.getCellStyle();
            XSSFCellStyle style = baseStyle.copy();
            style.setFont(boldFont);
            cell.setCellStyle(style);
        }

        // Add a border to the bottom row
        for (int i = 0; i < columns.length; i++) {
            XSSFCell cell = lastRow.getCell(i);
            XSSFCellStyle baseStyle = cell.getCellStyle();
            XSSFCellStyle style = baseStyle.copy();
            style.setBorderBottom(BorderStyle.THIN);
            cell.setCellStyle(style);
        }

        return currentRow;
    }
}
