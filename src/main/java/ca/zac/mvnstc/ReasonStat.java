/**
 * Insert extracted data to the result excel
 */
package ca.zac.mvnstc;

import java.text.SimpleDateFormat;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.*;

class ReasonStat extends StatBase {

    public ReasonStat(ArrayList<StockInfo> stockInfoList, Sheet reasonSheet, Integer numberOfReasons,
            CellStyle cellStyle) {
        super(stockInfoList, reasonSheet, numberOfReasons, cellStyle);
    }

    public ReasonStat(ArrayList<StockInfo> stockInfoList, Sheet reasonSheet, Integer numberOfReasons) {
        super(stockInfoList, reasonSheet, numberOfReasons);
    }

    @Override
    void insert() {
        // Loop excel and set the second column cells 0
        System.out.println("Rows in excel: " + reasonSheet.getLastRowNum());
        for (int j = 1; j <= reasonSheet.getLastRowNum(); j++) {
            currentRow = reasonSheet.getRow(j);
            if (currentRow == null) {
                continue;
            }
            if (currentRow.getCell(SECOND_COLUMN) == null) {
                continue;
            }
            currentRow.getCell(SECOND_COLUMN).setCellValue(0);
        }
        // Loop stock array list
        System.out.println("Stock count: " + stockInfoList.size());
        for (int i = 0; i < stockInfoList.size(); i++) {
            // Loop reasons
            for (int reasonIndex = 0; reasonIndex < stockInfoList.get(i).getReason().length; reasonIndex++) {
                Boolean isNewCategory = true;
                // Loop excel
                for (int j = 1; j <= reasonSheet.getLastRowNum() + 1; j++) {
                    // This is a blank row, a new row has to be create
                    if (isNewCategory && j == reasonSheet.getLastRowNum() + 1) {
                        Row newRow = reasonSheet.createRow(j);
                        // Write category name
                        newRow.createCell(CATEGORY_INDEX).setCellValue(stockInfoList.get(i).getReason()[reasonIndex]);
                        // write formated stock name with dates
                        Cell newCell = newRow.createCell(STOCK_LIST_INDEX);
                        newCell.setCellValue(formatStockInfo(i) + "\n");
                        newCell.setCellStyle(cellStyle);
                        // write the second column with 1
                        newRow.createCell(SECOND_COLUMN).setCellValue(1);
                        break;
                    }
                    // Compare existing category
                    currentRow = reasonSheet.getRow(j);
                    // blank row
                    if (currentRow == null) {
                        continue;
                    }
                    cellWithCategory = currentRow.getCell(CATEGORY_INDEX);
                    // Category can't be null
                    if (cellWithCategory == null
                            || !stockInfoList.get(i).getReason()[reasonIndex].contains(
                                    cellWithCategory.getStringCellValue().trim()))
                        continue;
                    // Found existing category
                    cellWithStockList = currentRow.getCell(STOCK_LIST_INDEX);
                    if (cellWithStockList == null) {
                        cellWithStockList = currentRow.createCell(STOCK_LIST_INDEX);
                    }
                    stockList = cellWithStockList.getStringCellValue() + formatStockInfo(i) + "\n";
                    cellWithStockList.setCellValue(stockList);
                    cellWithStockList.setCellStyle(cellStyle);
                    // update number of the second colunm
                    currentRow.getCell(SECOND_COLUMN)
                            .setCellValue(currentRow.getCell(SECOND_COLUMN).getNumericCellValue() + 1);
                    isNewCategory = false;
                    // Don't need to compare the following rows
                }
                // for first reason statistic
                if (numberOfReasons == 1) {
                    break;
                }

            }
        }
    }

    String formatStockInfo(int index) {
        SimpleDateFormat sdf = new SimpleDateFormat("hh:mm:ss");
        // Write increase dates that is greater than 1 at the end of each stock name
        String newStockInfo = stockInfoList.get(index).getName();
        if (newStockInfo.length() < 4) {
            // make the stock name to the same length
            newStockInfo += "    ";
        }
        newStockInfo = newStockInfo
                + "  " + sdf.format(stockInfoList.get(index).getFirstTime())
                + "  K" + stockInfoList.get(index).getOpenTimes()
                + "  " + sdf.format(stockInfoList.get(index).getLastTime())
                + "  " + parseAmountOn10Per(stockInfoList.get(index).getAmountOn10Per())
                + "/" + stockInfoList.get(index).getTotalAmount()
                + "  封" + stockInfoList.get(index).getSealBillAmount()
                + "  流" + stockInfoList.get(index).getSaleableShare();
        if (stockInfoList.get(index).getIncreaseDates() > 1) {
            newStockInfo = stockInfoList.get(index).getIncreaseDates().intValue() + newStockInfo;
        } else {
            newStockInfo = "  " + newStockInfo;
        }

        return newStockInfo;
    }

    Double parseAmountOn10Per(String s) {
        double multiplieer = 1.0;
        String input = s;
        if (input.endsWith("万")) {
            input = input.replace("万", "");
            multiplieer = 10000;
        } else if (input.endsWith("亿")) {
            input = input.replace("亿", "");
            multiplieer = 100000000;
        }
        try {
            Double value = Double.parseDouble(input) * multiplieer;
            return Math.round(value / 100000000 * 100.0) / 100.0;
        } catch (NumberFormatException e) {
            System.out.println(s + " dose not contain a number!");
        }
        return 0.00;
    }
}
