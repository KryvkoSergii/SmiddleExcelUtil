package ua.com.smiddle.excelutil;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import ua.com.smiddle.excelutil.exception.SEUConfigurerValidationException;
import ua.com.smiddle.excelutil.exception.SEUDataValidationException;

import java.io.IOException;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

/**
 * @author ksa on 14.03.17.
 * @project SmiddleCampaignManager
 */

public class ExcelExporter {
    private Map<Class, String> defaultTypePattern;
    private Map<Class, Short> defaultTypeFormat;
    private Configurer configurer;
    private Workbook wb;
    private final SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");


    //Constructors
    public ExcelExporter(Map<Class, String> defaultTypePattern, Configurer configurer) {
        this.defaultTypePattern = defaultTypePattern;
        this.configurer = configurer;
    }

    //Methods

//    public void exportAbonents(Collection<Abonent> abonentsList, Map<Integer, FieldMap> indexMapping, OutputStream os) throws Exception {
//        Workbook wb = new HSSFWorkbook();
//        String sheetName = "Реультат обзвона";
//        String reportName = null;
//        List<String[]> header = new ArrayList<>();
//        List<String[]> tableHeader = new ArrayList<>();
//
//        Configurer configurer = Configurer.buildNewConfigurer();
//        configurer.showReportName(true)
//                .showReportDetails(true)
//                .reportName("");
//        /**
//         * Contains headers of required abonent's data according to {@link FieldMap} of campaign.
//         */
//        List<String> importFields = indexMapping.entrySet().stream()
//                .map(i -> i.getValue().getFieldImport())
//                .collect(Collectors.toList());
//        /**
//         * container to store and transmit all useful data to excel render.
//         * Contains abonent's data and results.
//         */
//        List<String[]> data = new ArrayList<>();
//        /**
//         * Mapping of - column index in excel and {@link FieldMap} of campaign.
//         */
//        Set<Map.Entry<Integer, FieldMap>> indexMappingSet = indexMapping.entrySet();
//        int abonentMappingSize = indexMappingSet.size() - 1;
//        /**
//         * Max {@link ua.com.smiddle.SmiddleCampaignManager.core.model.ResultCode} column index counter in excel.
//         * Used for dynamical excel building.
//         */
//        int indexResultCounter = abonentMappingSize;
//        /**
//         * Mapping of - {@link ua.com.smiddle.SmiddleCampaignManager.core.model.ResultCode} code and {@link Wrapper},
//         * which contains name of {@link ua.com.smiddle.SmiddleCampaignManager.core.model.ResultCode} and it column index in excel.
//         * Name of {@link ua.com.smiddle.SmiddleCampaignManager.core.model.ResultCode} will be used as column title in next.
//         */
//        Map<String, Wrapper> resultCodeColumnIndex = new HashMap<>();
//        List<String> row;
//        for (Abonent a : abonentsList) {
//            row = new ArrayList<>();
//            /**
//             * Adding abonent's fields to temporary container {@code row}.
//             */
//            row.addAll(buildAbonentRow(indexMappingSet, a));
//            /**
//             * Adding results to temporary container {@code row}
//             */
//            indexResultCounter = appendResultRows(indexResultCounter, resultCodeColumnIndex, row, a);
//            /**
//             * converting to required data
//             */
//            data.add(row.toArray(new String[row.size()]));
//        }
//        /**
//         * Processing and retrieving headers of table
//         */
//        List<String> resultTitles = resultCodeColumnIndex.entrySet().stream()
//                .map(e -> e.getValue())
//                .sorted(Comparator.comparing(Wrapper::getIndex))
//                .map(e -> e.getName()).collect(Collectors.toList());
//        importFields.addAll(resultTitles);
//        tableHeader.add(importFields.toArray(new String[importFields.size()]));
//        /**
//         * rendering excel and writing to {@link OutputStream}
//         */
//        wb = useTemplate(wb, sheetName, reportName, header, tableHeader, data, 0, 0, false);
//        wb.write(os);
//    }


    //============================Services methods==============================


    public void bulidDocument(List<String[]> tableHeader, List<Object[]> data) throws Exception {
        if (defaultTypeFormat == null)
            throw new IllegalStateException("default type format are not initialized");
        Workbook wb = new HSSFWorkbook();
        validateConfigurer(configurer);
        validateTableHeadersSize(tableHeader,data,configurer);
        this.wb = useTemplate(wb, configurer.getSheetName(), configurer.getReportName(), configurer.getReportDetails(),
                tableHeader, data, configurer.getReportDateFrom(), configurer.getReportDateTo(), configurer.isShowReportDate());
    }

    public void writeDocument(OutputStream targetStream) throws IOException {
        wb.write(targetStream);
    }

    /**
     * validation of {@link Configurer} before build an document
     *
     * @param conf
     * @throws SEUConfigurerValidationException
     */
    private void validateConfigurer(Configurer conf) throws SEUConfigurerValidationException {
        if (conf.isShowReportName() && (conf.getReportName() == null || conf.getReportName().isEmpty()))
            throw new SEUConfigurerValidationException("Report name can't be empty");
        if (conf.isShowReportDetails() && (conf.getReportDetails() == null || conf.getReportDetails().isEmpty()))
            throw new SEUConfigurerValidationException("Report details can't be empty");
        if (conf.isShowReportDate() && (conf.getReportDateFrom() == 0 || conf.getReportDateTo() == 0))
            throw new SEUConfigurerValidationException("Report dateFrom and dateTo is not set");
    }

    /**
     * validate table headers, data and related date before build an document
     * @param tableHeader
     * @param data
     * @param conf
     * @throws SEUDataValidationException
     */
    private void validateTableHeadersSize(List<String[]> tableHeader, List<Object[]> data, Configurer conf) throws SEUDataValidationException {
        if (tableHeader == null)
            throw new SEUDataValidationException("tableHeader is not set");
        if (data == null)
            throw new SEUDataValidationException("data is not set");
        if (tableHeader.get(0).length != data.get(0).length)
            throw new SEUDataValidationException("Size of tableHeader[0]=" + tableHeader.get(0).length +
                    " not equal data[0]=" + data.get(0).length);
        if (conf.getCustomClassTypesRow() != null && tableHeader.get(0).length != conf.getCustomClassTypesRow().length)
            throw new SEUDataValidationException("Size of tableHeader[0]=" + tableHeader.get(0).length +
                    " not equal customClassTypesRow=" + conf.getCustomClassTypesRow().length);
        if (conf.getCustomPatternRow() != null && tableHeader.get(0).length != conf.getCustomPatternRow().length)
            throw new SEUDataValidationException("Size of tableHeader[0]=" + tableHeader.get(0).length +
                    " not equal customPatternRow=" + conf.getCustomPatternRow().length);
    }

    private Map<Integer, Short> buildCustomCellFormatter(List<String> pattersList) {
        final AtomicInteger i = new AtomicInteger(0);
        final CreationHelper createHelper = wb.getCreationHelper();
        Map<Integer, Short> format = pattersList.stream()
                .collect(Collectors.toMap(e -> i.getAndIncrement(), e -> createHelper.createDataFormat().getFormat(e)));
        return format;
    }

    private Workbook useTemplate(Workbook wb, String sheetName, String reportName, List<String[]> header,
                                 List<String[]> tableHeader, List<Object[]> data, long dateFrom, long dateTo, boolean dateRequired) {
        if (wb == null)
            throw new IllegalStateException("Workbook is not set");
        //Create new workbook and sheet
        Sheet sheet = wb.createSheet(WorkbookUtil.createSafeSheetName(sheetName));
        int offset = 0;
        //оглавление стилей
        CellStyle styleHeader = getHeaderStyle(wb);
        CellStyle tableHeaderStyle = getTableHeaderStyle(wb);
        CellStyle tableStyle = getTableStyle(wb);
        CellStyle generalHeaderStyle = getGeneralHeaderStyle(wb);
        CellStyle firstColumnStyle = getFirstColumnStyle(wb);
        //Определение границ
        short[] cellPosHeader = getMaxWidth(header);

        if (tableHeader != null) {
            short[] cellPosTable = getMaxWidth(tableHeader);
            if (cellPosHeader[0] > cellPosTable[0]) cellPosHeader[0] = cellPosTable[0];
            if (cellPosHeader[1] < cellPosTable[1]) cellPosHeader[1] = cellPosTable[1];
        }

        //отрисовка общего заголовка
        int rowNumber = 0;
        int headerLength = 0;
        Row row;
        Cell cell;
        for (rowNumber = 0; rowNumber < headerLength; rowNumber++) {
            row = sheet.createRow(rowNumber);
            for (int i = cellPosHeader[0]; i < cellPosHeader[1]; i++) {
                cell = row.createCell(i);
                cell.setCellStyle(generalHeaderStyle);
            }
        }
        //отрисовка заголовка
        if (reportName != null) {
            rowNumber++;
            row = sheet.createRow(rowNumber++);
            cell = row.createCell(1);
            cell.setCellValue(reportName);
            cell.setCellStyle(styleHeader);
        }
        //диапазон дат
        if (dateRequired) {
            rowNumber++;
            putDataRangeInReport(sheet, dateFrom, dateTo, rowNumber, firstColumnStyle);
            rowNumber++;
        }
        if (header != null && !header.isEmpty()) {
            sheet = appendData(sheet, header, offset, rowNumber, null, null, null, firstColumnStyle);
            rowNumber = sheet.getLastRowNum();
            rowNumber++;
        }
        //дата создания отчета
        if (dateRequired) {
            row = sheet.createRow(rowNumber++);
            cell = row.createCell(offset + 1);
            cell.setCellValue("Дата создания отчета: ");
            if (firstColumnStyle != null)
                cell.setCellStyle(firstColumnStyle);
            cell = row.createCell(offset + 2);
            cell.setCellValue(format.format(new Date()));
            rowNumber++;
        }
        if (tableHeader != null) {
            sheet = appendData(sheet, tableHeader, 0, rowNumber, tableHeaderStyle, null, null, null);
            rowNumber = sheet.getLastRowNum();
            rowNumber++;
            sheet.createFreezePane(cellPosHeader[1], rowNumber);
        }
        if (data != null) {
            appendData(sheet, data, 0, rowNumber, tableStyle, null, null, null);
            rowNumber = sheet.getLastRowNum();
        }
        //установка автоматического размера
        for (int i = 0; i < cellPosHeader[1]; i++)
            sheet.autoSizeColumn(i);
        return wb;
    }

    /**
     * Метод устанавливает диапазон дат в отчет при формировании *.xls формата.
     *
     * @param sheet    - {@link Sheet лист xls}
     * @param dateFrom - начальная дата
     * @param dateTo   - финальная дата
     * @return - {@link Sheet лист} с добавленной датой
     */
    private Sheet putDataRangeInReport(Sheet sheet, long dateFrom, long dateTo, int rowNumber, CellStyle firstColumn) {
        //диапазон дат
        Row row1 = sheet.createRow(rowNumber++);
        Cell c = row1.createCell(1);
        c.setCellValue("за период:");
        //установка стиля
        if (firstColumn != null)
            c.setCellStyle(firstColumn);
        //обработка даты
        if (dateFrom != 0 | dateTo != 0)
            row1.createCell(2).setCellValue(((dateFrom != 0) ? "c ".concat(format.format(new Date(dateFrom))) : "") +
                    ((dateTo != 0) ? " по ".concat(format.format(new Date(dateTo))) : ""));
        else
            row1.createCell(2).setCellValue("весь период");
        return sheet;
    }

    //CellStyles preparing
    private CellStyle getHeaderStyle(Workbook wb) {
        CellStyle styleHeader = wb.createCellStyle();
        styleHeader.setAlignment(CellStyle.ALIGN_CENTER);
        styleHeader.setVerticalAlignment(CellStyle.ALIGN_CENTER);
        Font font = wb.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        styleHeader.setFont(font);
        return styleHeader;
    }

    private CellStyle getTableHeaderStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.ALIGN_CENTER);
        style.setBorderTop(CellStyle.BORDER_MEDIUM);
        style.setBorderBottom(CellStyle.BORDER_MEDIUM);
        Font font = wb.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 10);
        style.setFont(font);
        return style;
    }

    private CellStyle getFirstColumnStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 10);
        style.setFont(font);
        return style;
    }

    private CellStyle getTableStyle(Workbook wb) {
        CellStyle styleThin = wb.createCellStyle();
        styleThin.setBorderTop(CellStyle.BORDER_THIN);
        styleThin.setBorderBottom(CellStyle.BORDER_THIN);
        styleThin.setBorderLeft(CellStyle.BORDER_THIN);
        styleThin.setBorderRight(CellStyle.BORDER_THIN);
        return styleThin;
    }

    private CellStyle getHeaderLowerStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setBorderBottom(CellStyle.BORDER_THIN);
        return style;
    }

    private CellStyle getHeaderUpperStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setBorderTop(CellStyle.BORDER_THIN);
        return style;
    }

    private CellStyle getGeneralHeaderStyle(Workbook wb) {
//        XSSFColor grey =new XSSFColor(new java.awt.Color(192,192,192));
//        cellStyle.setFillForegroundColor(grey);
        HSSFPalette palette = ((HSSFWorkbook) wb).getCustomPalette();
        // get the color which most closely matches the color you want to use
        //#59bab1
        HSSFColor myColor = palette.findSimilarColor(89, 186, 177);
        CellStyle style = wb.createCellStyle();
        style.setFillPattern(XSSFCellStyle.FINE_DOTS);
        style.setFillForegroundColor(myColor.getIndex());
        return style;
    }

    private short[] getFirstAndLastIndexInWidth(List<Row> header) {
        short[] result = new short[2];
        Row row;
        for (int i = 0; i < header.size(); i++) {
            row = header.get(i);
            if (i == 0) {
                result[0] = row.getFirstCellNum();
                result[1] = row.getLastCellNum();
            }
            if (result[0] > row.getFirstCellNum()) result[0] = row.getFirstCellNum();
            if (result[1] < row.getLastCellNum()) result[1] = row.getLastCellNum();
        }
        return result;
    }

    private short[] getMaxWidth(List<String[]> rows) {
        short[] result = new short[2];
        result[0] = 0;
        String[] row;
        for (int i = 0; i < rows.size(); i++) {
            row = rows.get(i);
            if (i == 0) result[1] = (short) row.length;
            if (result[1] < (short) row.length) result[1] = (short) row.length;
        }
        return result;
    }

    private Sheet appendData(Sheet sheet, List<Object[]> rows, int leftOffset, int beginRow, CellStyle style,
                             CellStyle upper, CellStyle lower, CellStyle firstColumn) {
        Row target;
        Object[] source;
        Cell cell;
        for (int i = 0; i < rows.size(); i++) {
            source = rows.get(i);
            target = sheet.createRow(beginRow++);
            for (int j = 0; j < source.length; j++) {
                Object cellValue = source[j];
                cell = target.createCell(leftOffset + j);
                // присвоение стиля
                if (style != null) cell.setCellStyle(style);
                //хедера и футера
                if (i == 0)
                    if (upper != null)
                        cell.setCellStyle(upper);
                if (i == rows.size() - 1)
                    if (lower != null)
                        cell.setCellStyle(lower);
                //присвоить стиль первой колонке
                if (j == 1)
                    if (firstColumn != null)
                        cell.setCellStyle(firstColumn);
                /** set value into cell */
                cell = setValueAndCellFormat(cell, cellValue);
            }

        }
        return sheet;
    }

    private Map<Class, Short> convertPatternToFormat(Map<Class, String> typePatterns) {
        final CreationHelper createHelper = wb.getCreationHelper();
        Map<Class, Short> formats = typePatterns.entrySet().stream()
                .collect(Collectors.toMap(k -> k.getKey(), v -> createHelper.createDataFormat().getFormat(v.getValue())));
        return formats;
    }

    private Cell setValueAndCellFormat(Cell cell, Object value, Short cellFormat, Class predefinedClassType) {
        /** set blank value to cell */
        if (value == null) {
            cell.setCellType(Cell.CELL_TYPE_BLANK);
            return cell;
        }

        cell = setCellTypeAndCellValueByValueType(cell, value);

        /** no predefinedClassType */
        if (predefinedClassType == null)
            return setCellFormat(cell, value, cellFormat, value.getClass());
        return setCellFormat(cell, value, cellFormat, predefinedClassType);
    }

    private Cell setCellFormat(Cell cell, Object value, Short cellFormat, Class predefinedClassType) {
        if (predefinedClassType == null)
            throw new IllegalArgumentException("Predefined class type is not set");
        if (predefinedClassType.getClass() != value.getClass())
            try {
                predefinedClassType.cast(value);
            } catch (ClassCastException e) {
                throw new IllegalArgumentException("Class type conflict " + predefinedClassType.getClass().getTypeName()
                        + " is not" + value.getClass().getTypeName());
            }
        Short format;
        /** use default cell format */
        if (cellFormat == null) format = defaultTypeFormat.get(predefinedClassType.getClass());
        /** use custom cell format */
        else format = cellFormat;
        cell.getCellStyle().setDataFormat(format);
        return cell;
    }

    /**
     * Method defines {@code value} classType and set cellType it to {@code cell}.
     *
     * @param cell  target cell, is type should be set
     * @param value object value, type of should be defined
     * @return target cell with type
     * @throws IllegalArgumentException type of {@code value} not supported.
     */
    private Cell setCellTypeAndCellValueByValueType(Cell cell, Object value) throws IllegalArgumentException {
        switch (value.getClass().getSimpleName()) {
            case "Long": {
                cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                cell.setCellValue((Long) value);
                break;
            }
            case "Integer": {
                cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                cell.setCellValue((Integer) value);
                break;
            }
            case "Short": {
                cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                cell.setCellValue((Byte) value);
                break;
            }
            case "Byte": {
                cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                cell.setCellValue((Byte) value);
                break;
            }
            case "Double": {
                cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                cell.setCellValue((Double) value);
                break;
            }
            case "Float": {
                cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                cell.setCellValue((Float) value);
                break;
            }
            case "Boolean": {
                cell.setCellType(Cell.CELL_TYPE_BOOLEAN);
                cell.setCellValue((Boolean) value);
                break;
            }
            case "String": {
                cell.setCellType(Cell.CELL_TYPE_STRING);
                cell.setCellValue((String) value);
                break;
            }
            case "Character": {
                cell.setCellType(Cell.CELL_TYPE_STRING);
                cell.setCellValue((String) value);
                break;
            }
            case "Date": {
                /** doesn't have self type, formatter should be set */
                break;
            }
            default:
                throw new IllegalArgumentException("Unsupported object value type");
        }
        return cell;
    }


    //InnerClasses
    private class Wrapper {
        private Integer index;
        private String name;


        //Constructor
        public Wrapper() {
        }

        public Wrapper(int index, String name) {
            this.index = index;
            this.name = name;
        }


        //Getters and setters
        public Integer getIndex() {
            return index;
        }

        public void setIndex(Integer index) {
            this.index = index;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }
    }

}
