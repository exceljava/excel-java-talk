/**
 * Example to show how Excel can be automated from a ribbon
 * toolbar using Jinx and com4j.
 *
 * This has the same functionality as the com4j-example
 * project but is invoked from a ribbon toolbar and
 * runs in Excel.
 *
 */
package com.exceljava.example;

import com.exceljava.com4j.JinxBridge;
import com.exceljava.com4j.excel.Range;
import com.exceljava.com4j.excel.XlDirection;
import com.exceljava.com4j.excel._Application;
import com.exceljava.com4j.office.IRibbonControl;
import com.exceljava.jinx.ExcelAction;
import com.exceljava.jinx.ExcelAddIn;
import com4j.Com4jObject;

public class TradeBlotter {
    private final ExcelAddIn addIn;

    public TradeBlotter(ExcelAddIn addIn) {
        this.addIn = addIn;
    }

    @ExcelAction
    public void writeTradeBlotter(IRibbonControl control) {
        _Application xl = JinxBridge.getApplication(this.addIn);

        Range header = xl.getRange("A1:E1");
        header.setValue(new String[] {
                "Ticker",
                "Quantity",
                "Price",
                "Status",
                "Error"
        });

        header.getInterior().setColor(0x643720);
        header.getFont().setColor(0xFFFFFF);
    }

    @ExcelAction
    public void readTradeBlotter(IRibbonControl control) {
        _Application xl = JinxBridge.getApplication(this.addIn);

        Range inputStart = xl.getRange("A2:E2");
        Range inputEnd = inputStart.getEnd(XlDirection.xlDown);
        Range inputRange = xl.getRange(inputStart, inputEnd);

        int numRows = inputRange.getRows().getCount();
        if (numRows > 100) {
            // If there are this many trades it's probably a mistake
            throw new RuntimeException("Too many trades");
        }

        for (int row = 0; row < inputRange.getRows().getCount(); row++) {
            // Get the cells for reporting the status and any errors
            Range statusCell = getCell(inputRange, row, 3);
            Range errorCell = getCell(inputRange, row, 4);

            // Get the input values for this row
            try {
                String ticker = (String)getCell(inputRange, row, 0).getValue();
                double quantity = (double)getCell(inputRange, row, 1).getValue();
                double price = (double)getCell(inputRange, row, 2).getValue();
                String status = (String)statusCell.getValue();

                // If the status is already "Booked" then there's nothing to do
                if (null != status && status.equals("Booked")) {
                    continue;
                }

                // Do some validation checks
                if (null == ticker) {
                    throw new RuntimeException("Ticker must be set");
                }

                if (price < 0) {
                    throw new RuntimeException("Price must be greater than 0");
                }

                if (quantity < 0) {
                    throw new RuntimeException("Quantity must be greater than 0");
                }

                // Enter the trade into the trade booking system...
                // TODO!

                // Mark the trade as booked
                statusCell.setValue("Booked");
                statusCell.getInterior().setColorIndex(0);

                // Clear the error cell
                errorCell.setValue("");
                errorCell.getInterior().setColorIndex(0);
            }
            catch (ClassCastException e) {
                statusCell.setValue("Error");
                statusCell.getInterior().setColor(0x00ffff);
                errorCell.setValue("Unexpected type");
                errorCell.getInterior().setColor(0x00ffff);
            }
            catch (NullPointerException e) {
                statusCell.setValue("Error");
                statusCell.getInterior().setColor(0x00ffff);
                errorCell.setValue("Missing value");
                errorCell.getInterior().setColor(0x00ffff);
            }
            catch (RuntimeException e) {
                statusCell.setValue("Error");
                statusCell.getInterior().setColor(0x00ffff);
                errorCell.setValue(null != e.getMessage() ? e.getMessage() : "Unknown error");
                errorCell.getInterior().setColor(0x00ffff);
            }
        }

        // Auto fit the status and error column
        Range errorCols = ((Com4jObject)xl.getColumns().getItem("D:E")).queryInterface(Range.class);
        errorCols.autoFit();
    }

    /**
     * Get a single cell from a Range using zero-based indexing.
     */
    private static Range getCell(Range range, int row, int col) {
        // Get the individual cell, remembering Excel uses 1 based indexing
        return ((Com4jObject)range.getItem(row + 1, col + 1)).queryInterface(Range.class);
    }
}
