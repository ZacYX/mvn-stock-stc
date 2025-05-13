/**
 * Stock name, increase reason, increase rate and dates in one excel row
 */
package ca.zac.mvnstc;

import java.util.Date;

class StockInfo implements Comparable<StockInfo> {
    String name; // 股票名称
    String detailedIndustry; // 细分行业
    String[] reason; // 涨停原因
    Double increaseRate; // 涨幅
    Double increaseDates; // 连涨天数

    Double totalAmount; // 总金额
    Double exchangeRate; // 换手
    String amountOn10Per; // 涨停成交额
    Double sealBillAmount; // 封单额
    Date firstTime; // 最终涨停时间
    Date lastTime; // 首次涨停时间
    Long openTimes; // 涨停开板次数
    Double saleableShare; // 流通市值

    static final String STOCK_NAME_HEADER = "名称";
    static final String STORCK_DETAILED_INDUSTRY_HEADER = "细分行业";
    static final String STOCK_REASON_HEADER = "涨停原因";
    static final String STOCK_INCREASE_RATE_HEADER = "涨幅";
    static final String STOCK_INCREASE_DATES_HEADER = "连续涨停天数";

    static final String TOTAL_AMOUNT = "总金额";
    static final String EXCHANGE_RATE = "换手";
    static final String AMOUNT_ON_10PER = "涨停成交额";
    static final String SEAL_BILL_AMOUNT = "封单额";
    static final String FIRST_TIME = "首次涨停时间";
    static final String OPEN_TIMES = "涨停开板次数";
    static final String LAST_TIME = "最终涨停时间";
    static final String SALEABLE_SHARE = "流通市值";

    static final String CELL_EMPTY_STRING = "--";
    static final Double STOCK_INCREASE_FLAG = 0.09;
    static final String STOCK_REASON_SPLITTER_REGEX_STRING = "\\+";

    @Override
    public int compareTo(StockInfo o) {
        return this.firstTime.compareTo(o.firstTime);
    }

    public String getName() {
        return this.name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getDetailedIndustry() {
        return this.detailedIndustry;
    }

    public void setDetailedIndustry(String detailedIndustry) {
        this.detailedIndustry = detailedIndustry;
    }

    public String[] getReason() {
        return this.reason;
    }

    public void setReason(String[] reason) {
        this.reason = reason;
    }

    public Double getIncreaseRate() {
        return this.increaseRate;
    }

    public void setIncreaseRate(Double increaseRate) {
        this.increaseRate = increaseRate;
    }

    public Double getIncreaseDates() {
        return this.increaseDates;
    }

    public void setIncreaseDates(Double increaseDates) {
        this.increaseDates = increaseDates;
    }

    public Double getTotalAmount() {
        return this.totalAmount;
    }

    public void setTotalAmount(Double totalAmount) {
        this.totalAmount = totalAmount;
    }

    public Double getExchangeRate() {
        return this.exchangeRate;
    }

    public void setExchangeRate(Double exchangeRate) {
        this.exchangeRate = exchangeRate;
    }

    public String getAmountOn10Per() {
        return this.amountOn10Per;
    }

    public void setAmountOn10Per(String amountOn10Per) {
        this.amountOn10Per = amountOn10Per;
    }

    public Double getSealBillAmount() {
        return this.sealBillAmount;
    }

    public void setSealBillAmount(Double sealBillAmount) {
        this.sealBillAmount = sealBillAmount;
    }

    public Date getFirstTime() {
        return this.firstTime;
    }

    public void setFirstTime(Date firstTime) {
        this.firstTime = firstTime;
    }

    public Date getLastTime() {
        return this.lastTime;
    }

    public void setLastTime(Date lastTime) {
        this.lastTime = lastTime;
    }

    public Long getOpenTimes() {
        return this.openTimes;
    }

    public void setOpenTimes(Long openTimes) {
        this.openTimes = openTimes;
    }

    public Double getSaleableShare() {
        return this.saleableShare;
    }

    public void setSaleableShare(Double saleableShare) {
        this.saleableShare = saleableShare;
    }

}
