package ca.zac.mvnstc;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOperator {
    static final String FIRST_REASON_SHEET_NAME = "概念首因";
    static final String ALL_REASON_SHEET_NAME = "概念全因";
    static final String ALL_REASON_COUNT_SHEET_NAME = "概念全因数字";
    static final String DETAILED_INDUSTRY_SHEET_NAME = "细分行业";
    static final String DETAILED_INDUSTRY_COUNT_SHEET_NAME = "细分行业数字";

    FileInputStream updaterFileInputStream;
    FileInputStream statisticResultFileInputStream;
    FileOutputStream newStatisticResultFileOutputStream;

    Workbook updaterWorkbook;
    Workbook statisticResultWorkbook;

    CellStyle cellStyle;

    Sheet updaterSheet;
    Sheet firstReasonSheet;
    Sheet allReasonSheet;
    Sheet allReasonCountSheet;
    Sheet detailedIndustrySheet;
    Sheet detailedIndustryCountSheet;

    public ExcelOperator(String updaterFilePath, String statisticResultFilePath, String newStatisticResultFilePat)
            throws IOException {
        try {
            this.updaterFileInputStream = new FileInputStream(updaterFilePath);
            updaterWorkbook = new XSSFWorkbook(this.updaterFileInputStream);
            updaterSheet = (Sheet) updaterWorkbook.getSheetAt(0);
        } catch (IOException e) {
            System.out.println("Open updater failed!");
            throw e;
        }
        try {
            this.statisticResultFileInputStream = new FileInputStream(statisticResultFilePath);
            statisticResultWorkbook = new XSSFWorkbook(this.statisticResultFileInputStream);
        } catch (IOException e) {
            System.out.println("marketinfo file not found");
            statisticResultWorkbook = new XSSFWorkbook();
        }
        cellStyle = statisticResultWorkbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.LEFT); // 水平居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
        cellStyle.setWrapText(true);
        try {
            Date date = new Date();
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MMdd-hhmm");
            this.newStatisticResultFileOutputStream = new FileOutputStream(
                    newStatisticResultFilePat + simpleDateFormat.format(date) + "-marketinfo.xlsx");

            firstReasonSheet = (Sheet) statisticResultWorkbook.getSheet(FIRST_REASON_SHEET_NAME);
            if (firstReasonSheet == null) {
                firstReasonSheet = (Sheet) statisticResultWorkbook.createSheet(FIRST_REASON_SHEET_NAME);
            }
            allReasonSheet = (Sheet) statisticResultWorkbook.getSheet(ALL_REASON_SHEET_NAME);
            if (allReasonSheet == null) {
                allReasonSheet = (Sheet) statisticResultWorkbook.createSheet(ALL_REASON_SHEET_NAME);
            }
            allReasonCountSheet = (Sheet) statisticResultWorkbook.getSheet(ALL_REASON_COUNT_SHEET_NAME);
            if (allReasonCountSheet == null) {
                allReasonCountSheet = (Sheet) statisticResultWorkbook.createSheet(ALL_REASON_COUNT_SHEET_NAME);
            }
            detailedIndustrySheet = (Sheet) statisticResultWorkbook.getSheet(DETAILED_INDUSTRY_SHEET_NAME);
            if (detailedIndustrySheet == null) {
                detailedIndustrySheet = (Sheet) statisticResultWorkbook.createSheet(DETAILED_INDUSTRY_SHEET_NAME);
            }
            detailedIndustryCountSheet = (Sheet) statisticResultWorkbook.getSheet(DETAILED_INDUSTRY_COUNT_SHEET_NAME);
            if (detailedIndustryCountSheet == null) {
                detailedIndustryCountSheet = (Sheet) statisticResultWorkbook
                        .createSheet(DETAILED_INDUSTRY_COUNT_SHEET_NAME);
            }
        } catch (Exception e) {
            System.out.println("Exception in ExcelOperator");
            e.printStackTrace();
        }

    }

    public void close() {
        try {
            statisticResultWorkbook.write(newStatisticResultFileOutputStream);
            updaterWorkbook.close();
            statisticResultWorkbook.close();
            updaterFileInputStream.close();
            statisticResultFileInputStream.close();
            newStatisticResultFileOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public FileInputStream getUpdaterFileInputStream() {
        return this.updaterFileInputStream;
    }

    public void setUpdaterFileInputStream(FileInputStream updaterFileInputStream) {
        this.updaterFileInputStream = updaterFileInputStream;
    }

    public FileInputStream getStatisticResultFileInputStream() {
        return this.statisticResultFileInputStream;
    }

    public void setStatisticResultFileInputStream(FileInputStream statisticResultFileInputStream) {
        this.statisticResultFileInputStream = statisticResultFileInputStream;
    }

    public FileOutputStream getNewStatisticResultFileOutputStream() {
        return this.newStatisticResultFileOutputStream;
    }

    public void setNewStatisticResultFileOutputStream(FileOutputStream newStatisticResultFileOutputStream) {
        this.newStatisticResultFileOutputStream = newStatisticResultFileOutputStream;
    }

    public Workbook getUpdaterWorkbook() {
        return this.updaterWorkbook;
    }

    public void setUpdaterWorkbook(Workbook updaterWorkbook) {
        this.updaterWorkbook = updaterWorkbook;
    }

    public Workbook getStatisticResultWorkbook() {
        return this.statisticResultWorkbook;
    }

    public void setStatisticResultWorkbook(Workbook statisticResultWorkbook) {
        this.statisticResultWorkbook = statisticResultWorkbook;
    }

    public CellStyle getCellStyle() {
        return this.cellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public Sheet getUpdaterSheet() {
        return this.updaterSheet;
    }

    public void setUpdaterSheet(Sheet updaterSheet) {
        this.updaterSheet = updaterSheet;
    }

    public Sheet getFirstReasonSheet() {
        return this.firstReasonSheet;
    }

    public void setFirstReasonSheet(Sheet firstReasonSheet) {
        this.firstReasonSheet = firstReasonSheet;
    }

    public Sheet getAllReasonSheet() {
        return this.allReasonSheet;
    }

    public void setAllReasonSheet(Sheet allReasonSheet) {
        this.allReasonSheet = allReasonSheet;
    }

    public Sheet getAllReasonCountSheet() {
        return this.allReasonCountSheet;
    }

    public void setAllReasonCountSheet(Sheet allReasonCountSheet) {
        this.allReasonCountSheet = allReasonCountSheet;
    }

    public Sheet getDetailedIndustrySheet() {
        return this.detailedIndustrySheet;
    }

    public void setDetailedIndustrySheet(Sheet detailedIndustrySheet) {
        this.detailedIndustrySheet = detailedIndustrySheet;
    }

    public Sheet getDetailedIndustryCountSheet() {
        return this.detailedIndustryCountSheet;
    }

    public void setDetailedIndustryCountSheet(Sheet detailedIndustryCountSheet) {
        this.detailedIndustryCountSheet = detailedIndustryCountSheet;
    }

}
