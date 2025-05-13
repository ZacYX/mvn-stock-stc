/**
 * Extract stock infomation from the downloaded excel 
 * file from ths.
 */
package ca.zac.mvnstc;

import java.util.ArrayList;
import java.util.Collections;

import org.apache.poi.ss.usermodel.*;

class Updater {
    int nameColumnIndex;
    int detailedIndustryColumnIndex; // 细分行业
    int reasonColumnIndex; // 涨停原因
    int increaseRateColumnIndex; // 涨幅
    int increaseDatesColumnIndex; // 连续涨停天数

    int totalAmountIndex; // 总金额
    int exchangeRateIndex; // 换手
    int amountOn10PerIndex; // 涨停成交额
    int sealBillAmountIndex; // 封单额
    int firstTimeIndex; // 最终涨停时间
    int openTimesIndex; // 涨停开板次数
    int lastTimeIndex; // 首次涨停时间
    int saleableShareIndex; // 流通市值

    Sheet updaterSheet;
    Row currentRow;
    Cell cellWithName;
    Cell cellWithDetailedIndustry;
    Cell cellWithReason;
    Cell cellWithIncreaseRate;
    Cell cellWithIncreaseDates;

    Cell cellWithTotalAmount;
    Cell cellWithExchageRate;
    Cell cellWithAmountOn10Per;
    Cell cellWithSealBillAmount;
    Cell cellWithFirstTime;
    Cell cellWithLastTime;
    Cell cellWithOpenTimes;
    Cell cellWithSaleableShare;

    StockInfo stockInfo;
    ArrayList<StockInfo> stockInfoList;

    public Updater(Sheet updaterSheet) {
        this.updaterSheet = updaterSheet;
        this.nameColumnIndex = -1;
        this.detailedIndustryColumnIndex = -1;
        this.reasonColumnIndex = -1;
        this.increaseDatesColumnIndex = -1;
        this.increaseRateColumnIndex = -1;

        this.totalAmountIndex = -1; // 总金额
        this.exchangeRateIndex = -1; // 换手
        this.amountOn10PerIndex = -1; // 涨停成交额
        this.sealBillAmountIndex = -1; // 封单额
        this.firstTimeIndex = -1; // 最终涨停时间
        this.openTimesIndex = -1; // 涨停开板次数
        this.lastTimeIndex = -1; // 首次涨停时间
        this.saleableShareIndex = -1;

    }

    public ArrayList<StockInfo> getData() {
        this.prepare();
        this.process();
        Collections.sort(this.stockInfoList);
        return this.stockInfoList;
    }

    // prepare workbook, worksheet, collumn index of name, reason, increase rate and
    // dates
    void prepare() {
        this.stockInfo = new StockInfo();
        this.stockInfoList = new ArrayList<StockInfo>();
        Row headerRow = this.updaterSheet.getRow(0); // First row
        for (Cell cell : headerRow) {
            // There is a space before this string
            if (cell.getStringCellValue().trim().contains(StockInfo.STOCK_NAME_HEADER)) {
                this.nameColumnIndex = cell.getColumnIndex();
            }
            if (cell.getStringCellValue().trim().contains(StockInfo.STORCK_DETAILED_INDUSTRY_HEADER)) {
                this.detailedIndustryColumnIndex = cell.getColumnIndex();
            }
            if (cell.getStringCellValue().trim().contains(StockInfo.STOCK_REASON_HEADER)) {
                this.reasonColumnIndex = cell.getColumnIndex();
            }
            // More than one contain this string
            if (cell.getStringCellValue().trim().equals(StockInfo.STOCK_INCREASE_RATE_HEADER)) {
                this.increaseRateColumnIndex = cell.getColumnIndex();
            }
            if (cell.getStringCellValue().trim().contains(StockInfo.STOCK_INCREASE_DATES_HEADER)) {
                this.increaseDatesColumnIndex = cell.getColumnIndex();
            }
            // Added on May 12
            if (cell.getStringCellValue().trim().contains(StockInfo.TOTAL_AMOUNT)) {
                this.totalAmountIndex = cell.getColumnIndex();
            }
            if (cell.getStringCellValue().trim().equals(StockInfo.EXCHANGE_RATE)) {
                this.exchangeRateIndex = cell.getColumnIndex();
            }
            if (cell.getStringCellValue().trim().contains(StockInfo.AMOUNT_ON_10PER)) {
                this.amountOn10PerIndex = cell.getColumnIndex();
            }
            if (cell.getStringCellValue().trim().equals(StockInfo.SEAL_BILL_AMOUNT)) {
                this.sealBillAmountIndex = cell.getColumnIndex();
            }
            if (cell.getStringCellValue().trim().contains(StockInfo.FIRST_TIME)) {
                this.firstTimeIndex = cell.getColumnIndex();
            }
            if (cell.getStringCellValue().trim().contains(StockInfo.OPEN_TIMES)) {
                this.openTimesIndex = cell.getColumnIndex();
            }
            if (cell.getStringCellValue().trim().contains(StockInfo.LAST_TIME)) {
                this.lastTimeIndex = cell.getColumnIndex();
            }
            if (cell.getStringCellValue().trim().contains(StockInfo.SALEABLE_SHARE)) {
                this.saleableShareIndex = cell.getColumnIndex();
            }
            // Stop loop after getting all index
            if (this.nameColumnIndex != -1 && this.reasonColumnIndex != -1
                    && this.increaseRateColumnIndex != -1
                    && this.increaseDatesColumnIndex != -1
                    && this.detailedIndustryColumnIndex != -1) {
                break;
            }
        }
        System.out.println("Index:  name, " + nameColumnIndex
                + "    Exchange Rate, " + exchangeRateIndex
                + "    Detailed_industry, " + detailedIndustryColumnIndex
                + "    Reason, " + reasonColumnIndex
                + "    Rate, " + increaseRateColumnIndex
                + "    Saleable, " + saleableShareIndex
                + "    dates, " + increaseDatesColumnIndex);
        System.out.println(
                "Total amount, " + totalAmountIndex
                        + "    Amount on 10%, " + amountOn10PerIndex
                        + "    Seal Bill, " + sealBillAmountIndex
                        + "    First Time, " + firstTimeIndex
                        + "    Open times, " + openTimesIndex
                        + "    Last Time, " + lastTimeIndex);
    }

    // Get stock name, reason, increase rate, increase dates according row index
    void process() {
        for (int i = 1; i < this.updaterSheet.getLastRowNum(); i++) {
            this.currentRow = this.updaterSheet.getRow(i);
            this.cellWithName = this.currentRow.getCell(nameColumnIndex);
            this.cellWithDetailedIndustry = this.currentRow.getCell(detailedIndustryColumnIndex);
            this.cellWithReason = this.currentRow.getCell(reasonColumnIndex);
            this.cellWithIncreaseRate = this.currentRow.getCell(increaseRateColumnIndex);
            this.cellWithIncreaseDates = this.currentRow.getCell(increaseDatesColumnIndex);

            this.cellWithTotalAmount = this.currentRow.getCell(totalAmountIndex);
            this.cellWithExchageRate = this.currentRow.getCell(exchangeRateIndex);
            this.cellWithAmountOn10Per = this.currentRow.getCell(amountOn10PerIndex);
            this.cellWithSealBillAmount = this.currentRow.getCell(sealBillAmountIndex);
            this.cellWithFirstTime = this.currentRow.getCell(firstTimeIndex);
            this.cellWithOpenTimes = this.currentRow.getCell(openTimesIndex);
            this.cellWithLastTime = this.currentRow.getCell(lastTimeIndex);
            this.cellWithSaleableShare = this.currentRow.getCell(saleableShareIndex);
            // read cells' content
            try {
                // Name
                this.stockInfo.setName(this.cellWithName.getStringCellValue().trim());
                // Detailed industry
                this.stockInfo.setDetailedIndustry(this.cellWithDetailedIndustry.getStringCellValue().trim());
                // Reason: ****+*****+*****+****, get the one before the first "+", or "--"
                this.stockInfo.setReason(this.cellWithReason.getStringCellValue().trim()
                        .split(StockInfo.STOCK_REASON_SPLITTER_REGEX_STRING));
                // Increase rate
                if (this.cellWithIncreaseRate.getCellType() == CellType.NUMERIC) {
                    this.stockInfo.setIncreaseRate(this.cellWithIncreaseRate.getNumericCellValue());
                }
                // Increase dates
                if (this.cellWithIncreaseDates.getCellType() == CellType.NUMERIC) {
                    this.stockInfo.setIncreaseDates(this.cellWithIncreaseDates.getNumericCellValue());
                }

                // Added on May 13
                if (this.cellWithTotalAmount.getCellType() == CellType.NUMERIC) {
                    // 金额单位：亿
                    this.stockInfo.setTotalAmount(toYi(this.cellWithTotalAmount.getNumericCellValue()));
                }
                if (cellWithExchageRate.getCellType() == CellType.NUMERIC) {
                    stockInfo.setExchangeRate(cellWithExchageRate.getNumericCellValue());
                }
                if (cellWithAmountOn10Per.getCellType() == CellType.STRING) {
                    stockInfo.setAmountOn10Per(cellWithAmountOn10Per.getStringCellValue());
                }
                if (cellWithSealBillAmount.getCellType() == CellType.NUMERIC) {
                    stockInfo.setSealBillAmount(toYi(cellWithSealBillAmount.getNumericCellValue()));
                }
                if (cellWithFirstTime.getCellType() == CellType.NUMERIC
                        && DateUtil.isCellDateFormatted(cellWithFirstTime)) {
                    stockInfo.setFirstTime((cellWithFirstTime.getDateCellValue()));
                }
                if (cellWithLastTime.getCellType() == CellType.NUMERIC
                        && DateUtil.isCellDateFormatted(cellWithLastTime)) {
                    stockInfo.setLastTime((cellWithLastTime.getDateCellValue()));
                }
                if (cellWithOpenTimes.getCellType() == CellType.NUMERIC) {
                    stockInfo.setOpenTimes(Math.round(cellWithOpenTimes.getNumericCellValue()));
                }
                if (cellWithSaleableShare.getCellType() == CellType.NUMERIC) {
                    stockInfo.setSaleableShare(toYi(cellWithSaleableShare.getNumericCellValue()));
                }

                // System.out.println("name, " + stockInfo.getName()
                // + " Detailed_industry, " + stockInfo.getDetailedIndustry()
                // + " Reason, " + stockInfo.getReason()[0]
                // + " Rate, " + stockInfo.getIncreaseRate()
                // + " dates, " + stockInfo.getIncreaseDates());

                // Not "--" and increase rate > 0.09 means a valid info, and add it to the list
                if (
                // !this.stockInfo.getReason()[0].equals(StockInfo.CELL_EMPTY_STRING)
                this.stockInfo.getIncreaseRate() > StockInfo.STOCK_INCREASE_FLAG
                        && this.stockInfo.getName().length() > 0
                        && this.stockInfo.getIncreaseDates() > 0) {
                    this.stockInfoList.add(this.stockInfo);
                    this.stockInfo = new StockInfo();
                } else {
                    /*
                     * System.out.println(this.stockInfo.getName() + "    "
                     * + this.stockInfo.getReason()[0] + "    "
                     * + this.stockInfo.getIncreaseRate() + "    "
                     * + this.stockInfo.getIncreaseDates());
                     */
                }
            } catch (Exception e) {
                e.printStackTrace();
                System.out.println("Exception line: " + i + "  ");
                continue;
            }
        }
        System.out.println("Total items: " + stockInfoList.size());
    }

    Double toYi(Double num) {
        return Math.round(num / 100000000 * 100.0) / 100.0;
    }
}