package ua.com.smiddle.excelutil;

import ua.com.smiddle.excelutil.model.CellStylePolicy;
import ua.com.smiddle.excelutil.model.ExceptionProcessingPolicy;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by ksa on 26.03.17.
 */
public class Configurer {
    private boolean showReportName;
    private boolean showReportDetails;
    private boolean showReportDate;
    private long reportDateFrom;
    private long reportDateTo;
    private String reportName;
    private List<Object[]> reportDetails;
    private String sheetName;
    private Class[] customClassTypesRow;
    private String[] customPatternRow;
    private ExceptionProcessingPolicy exceptionProcessingPolicy;
    private CellStylePolicy cellStylePolicy;


    //Constructors
    private Configurer() {
    }

    private Configurer(boolean showReportName, boolean showReportDetails, boolean showReportDate, String reportName,
                       List<Object[]> reportDetails, String sheetName, long reportDateFrom, long reportDateTo,
                       Class[] customClassTypesRow, String[] customPatternRow, ExceptionProcessingPolicy exceptionProcessingPolicy
            , CellStylePolicy cellStylePolicy) {
        this.showReportName = showReportName;
        this.showReportDetails = showReportDetails;
        this.showReportDate = showReportDate;
        this.reportName = reportName;
        this.reportDetails = reportDetails;
        this.sheetName = sheetName;
        this.reportDateFrom = reportDateFrom;
        this.reportDateTo = reportDateTo;
        this.customClassTypesRow = customClassTypesRow;
        this.customPatternRow = customPatternRow;
        this.exceptionProcessingPolicy = exceptionProcessingPolicy;
        this.cellStylePolicy = cellStylePolicy;
    }


    //Methods
    public static Configurer buildNewConfigurer() {
        return new Configurer(false, false, false, "REPORT",
                new ArrayList<>(), "Sheet1", 0L, 0L, null, null,
                ExceptionProcessingPolicy.WRITE_EXCEPTION_TO_CELL, CellStylePolicy.COLUMN);
    }

    public Configurer showReportName(boolean showReportName) {
        this.showReportName = showReportName;
        return this;
    }

    public Configurer showReportDetails(boolean showReportDetails) {
        this.showReportDetails = showReportDetails;
        return this;
    }

    public Configurer showReportDate(boolean showReportDate) {
        this.showReportDate = showReportDate;
        return this;
    }

    public Configurer reportName(String reportName) {
        this.showReportName = true;
        this.reportName = reportName;
        return this;
    }

    public Configurer reportDetails(List<Object[]> reportDetails) {
        this.showReportDetails = true;
        this.reportDetails = reportDetails;
        return this;
    }

    public Configurer sheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    public Configurer reportDateFrom(long reportDateFrom) {
        this.reportDateFrom = reportDateFrom;
        return this;
    }

    public Configurer reportDateTo(long reportDateTo) {
        this.reportDateTo = reportDateTo;
        return this;
    }

    public Configurer customClassTypesRow(Class[] customClassTypesRow) {
        this.customClassTypesRow = customClassTypesRow;
        return this;
    }

    public Configurer customPatternRow(String[] customPatternRow) {
        this.customPatternRow = customPatternRow;
        return this;
    }

    public Configurer exceptionProcessingPolicy(ExceptionProcessingPolicy exceptionProcessingPolicy) {
        this.exceptionProcessingPolicy = exceptionProcessingPolicy;
        return this;
    }

    public Configurer cellStylePolicy(CellStylePolicy stylePolicy) {
        this.cellStylePolicy = stylePolicy;
        return this;
    }

    public ExceptionProcessingPolicy getExceptionProcessingPolicy() {
        return exceptionProcessingPolicy;
    }

    public CellStylePolicy getCellStylePolicy() {
        return cellStylePolicy;
    }

    public long getReportDateTo() {
        return reportDateTo;
    }

    public long getReportDateFrom() {
        return reportDateFrom;
    }

    public boolean isShowReportName() {
        return showReportName;
    }

    public boolean isShowReportDetails() {
        return showReportDetails;
    }

    public boolean isShowReportDate() {
        return showReportDate;
    }

    public String getReportName() {
        return reportName;
    }

    public List<Object[]> getReportDetails() {
        return reportDetails;
    }

    public String getSheetName() {
        return sheetName;
    }

    public Class[] getCustomClassTypesRow() {
        return customClassTypesRow;
    }

    public String[] getCustomPatternRow() {
        return customPatternRow;
    }

    public Configurer getConfigurer() {
        return this;
    }
}
