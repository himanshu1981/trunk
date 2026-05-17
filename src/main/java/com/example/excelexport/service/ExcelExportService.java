package com.example.excelexport.service;

import com.example.excelexport.model.*;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;

import java.io.OutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * Core Excel export service.
 *
 * Uses Apache POI SXSSF (Streaming Usermodel) which writes rows to a temp file as
 * they are produced, keeping only a sliding window of rows in heap memory.
 * This makes exports of 1 000 000+ rows feasible with a fixed, small heap footprint.
 *
 * Typical call:
 * <pre>
 *   excelExportService.exportToStream(config, dataProvider, httpServletResponse.getOutputStream());
 * </pre>
 */
@Slf4j
@Service
public class ExcelExportService {

    // Default: corporate dark blue
    private static final int[] DEFAULT_HEADER_BG = {31, 73, 125};

    private static final short HEADER_FONT_SIZE = 11;
    private static final short DATA_FONT_SIZE   = 10;

    // -----------------------------------------------------------------------
    // Public API
    // -----------------------------------------------------------------------

    /**
     * Streams an .xlsx file to {@code outputStream}.
     *
     * @param config       columns, freeze settings, styling options
     * @param dataProvider source that pushes rows via the ExcelRowWriter callback
     * @param outputStream destination — typically HttpServletResponse.getOutputStream()
     */
    public void exportToStream(ExcelExportConfig config,
                               ExcelDataProvider dataProvider,
                               OutputStream outputStream) throws Exception {

        int windowSize = config.getRowAccessWindowSize() > 0
                ? config.getRowAccessWindowSize() : 1_000;

        log.info("Excel export starting — sheet='{}', freezeCols={}, freezeRows={}, windowSize={}",
                config.getSheetName(), config.getFreezeColumnCount(),
                config.getFreezeRowCount(), windowSize);

        // SXSSFWorkbook spills rows beyond windowSize to a temp file automatically.
        // setCompressTempFiles(true) uses GZIP for those temp files, trading CPU for disk.
        try (SXSSFWorkbook workbook = new SXSSFWorkbook(windowSize)) {
            workbook.setCompressTempFiles(true);

            String sheetName = StringUtils.hasText(config.getSheetName())
                    ? config.getSheetName() : "Sheet1";
            SXSSFSheet sheet = workbook.createSheet(sheetName);
            sheet.setDefaultRowHeightInPoints(15);

            List<ExcelColumnConfig> columns = config.getColumns();
            int colCount = columns.size();

            // ----------------------------------------------------------------
            // Pre-create ALL cell styles BEFORE any rows are written.
            // POI has a hard limit of 64 000 styles per workbook; creating one
            // per cell would exhaust it instantly for large exports.
            // ----------------------------------------------------------------
            CellStyle            headerStyle   = buildHeaderStyle(workbook, config);
            Map<Integer, CellStyle> colStyles  = buildColumnStyles(workbook, columns);

            // ----------------------------------------------------------------
            // Header row (index 0)
            // ----------------------------------------------------------------
            Row headerRow = sheet.createRow(0);
            headerRow.setHeightInPoints(22);
            for (int i = 0; i < colCount; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columns.get(i).getHeader());
                cell.setCellStyle(headerStyle);
            }

            // ----------------------------------------------------------------
            // Freeze pane — must be set after the sheet is created but the
            // position is independent of how many rows have been written.
            // createFreezePane(colSplit, rowSplit):
            //   colSplit = first UNfrozen column (0-based)
            //   rowSplit = first UNfrozen row    (0-based)
            // ----------------------------------------------------------------
            int freezeCols = Math.max(config.getFreezeColumnCount(), 0);
            int freezeRows = config.getFreezeRowCount() > 0 ? config.getFreezeRowCount() : 1;
            sheet.createFreezePane(freezeCols, freezeRows);

            // ----------------------------------------------------------------
            // Auto-filter on the header row — spans all columns
            // ----------------------------------------------------------------
            sheet.setAutoFilter(new CellRangeAddress(0, 0, 0, colCount - 1));

            // ----------------------------------------------------------------
            // Stream data rows — dataProvider calls rowWriter once per record
            // ----------------------------------------------------------------
            AtomicInteger rowNum = new AtomicInteger(1);

            dataProvider.fetchData(rowData -> {
                Row row = sheet.createRow(rowNum.getAndIncrement());
                for (int i = 0; i < colCount; i++) {
                    ExcelColumnConfig colCfg = columns.get(i);
                    Cell cell = row.createCell(i);
                    cell.setCellStyle(colStyles.get(i));
                    setCellValue(cell, rowData.get(colCfg.getFieldName()), colCfg.getDataType());
                }
            });

            int dataRows = rowNum.get() - 1;
            log.info("Excel export — {} data rows written, setting column widths", dataRows);

            // ----------------------------------------------------------------
            // Column widths — must be set AFTER rows are written when estimating,
            // but since SXSSF flushes rows we use a static estimate rather than
            // autoSizeColumn (which only works for rows still in the window).
            // ----------------------------------------------------------------
            for (int i = 0; i < colCount; i++) {
                ExcelColumnConfig col = columns.get(i);
                int width = col.getWidthChars() > 0
                        ? col.getWidthChars() * 256
                        : estimateWidth(col);
                sheet.setColumnWidth(i, Math.min(width, 255 * 256)); // POI max = 255 chars
            }

            workbook.write(outputStream);
            outputStream.flush();
            // workbook.close() calls dispose() which deletes the SXSSF temp files
        }

        log.info("Excel export complete");
    }

    // -----------------------------------------------------------------------
    // Style builders
    // -----------------------------------------------------------------------

    private CellStyle buildHeaderStyle(SXSSFWorkbook workbook, ExcelExportConfig config) {
        // Cast is safe: SXSSF delegates style creation to the underlying XSSFWorkbook
        XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();

        int[] rgb = config.getHeaderBgRgb() != null ? config.getHeaderBgRgb() : DEFAULT_HEADER_BG;
        byte[] rgbBytes = {(byte) rgb[0], (byte) rgb[1], (byte) rgb[2]};
        style.setFillForegroundColor(new XSSFColor(rgbBytes, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        font.setFontHeightInPoints(HEADER_FONT_SIZE);
        style.setFont(font);

        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);

        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    /**
     * Creates one CellStyle per column (not per cell).
     * Columns sharing the same type+format share a Font object to stay under the style limit.
     */
    private Map<Integer, CellStyle> buildColumnStyles(SXSSFWorkbook workbook,
                                                       List<ExcelColumnConfig> columns) {
        Map<Integer, CellStyle> styles = new HashMap<>();
        DataFormat dataFormat = workbook.createDataFormat();

        // Single shared font for all data cells
        Font dataFont = workbook.createFont();
        dataFont.setFontHeightInPoints(DATA_FONT_SIZE);

        for (int i = 0; i < columns.size(); i++) {
            ExcelColumnConfig col = columns.get(i);
            CellStyle style = workbook.createCellStyle();
            style.setFont(dataFont);

            // International standard: numbers / dates right-aligned, text left-aligned
            style.setAlignment(alignmentFor(col.getDataType()));
            style.setVerticalAlignment(VerticalAlignment.CENTER);

            String fmt = StringUtils.hasText(col.getNumberFormat())
                    ? col.getNumberFormat()
                    : defaultFormatFor(col.getDataType());
            if (fmt != null) {
                style.setDataFormat(dataFormat.getFormat(fmt));
            }

            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);

            styles.put(i, style);
        }
        return styles;
    }

    // -----------------------------------------------------------------------
    // Alignment & format helpers
    // -----------------------------------------------------------------------

    private HorizontalAlignment alignmentFor(ColumnDataType type) {
        return switch (type) {
            case NUMERIC, INTEGER, CURRENCY, PERCENTAGE, DATE, DATETIME -> HorizontalAlignment.RIGHT;
            case STRING, ALPHANUMERIC -> HorizontalAlignment.LEFT;
        };
    }

    private String defaultFormatFor(ColumnDataType type) {
        return switch (type) {
            case NUMERIC     -> "#,##0.##";
            case INTEGER     -> "#,##0";
            case CURRENCY    -> "$#,##0.00";
            case PERCENTAGE  -> "0.00%";
            case DATE        -> "yyyy-mm-dd";
            case DATETIME    -> "yyyy-mm-dd hh:mm:ss";
            case STRING, ALPHANUMERIC -> null; // General / @
        };
    }

    // -----------------------------------------------------------------------
    // Cell value writer — maps JDBC types to POI cell types
    // -----------------------------------------------------------------------

    private void setCellValue(Cell cell, Object value, ColumnDataType type) {
        if (value == null) {
            cell.setBlank();
            return;
        }

        switch (type) {
            case NUMERIC, CURRENCY -> {
                if (value instanceof Number n) cell.setCellValue(n.doubleValue());
                else tryDouble(cell, value.toString());
            }
            case INTEGER -> {
                if (value instanceof Number n) cell.setCellValue(n.longValue());
                else tryLong(cell, value.toString());
            }
            // Percentage values are stored as 0-100 in the DB; Excel 0.00% format expects 0-1
            case PERCENTAGE -> {
                if (value instanceof Number n) cell.setCellValue(n.doubleValue() / 100.0);
                else tryDoubleDiv100(cell, value.toString());
            }
            case DATE -> {
                if      (value instanceof LocalDate ld)         cell.setCellValue(ld);
                else if (value instanceof java.sql.Date sd)     cell.setCellValue(sd.toLocalDate());
                else if (value instanceof java.util.Date d)     cell.setCellValue(d);
                else                                            cell.setCellValue(value.toString());
            }
            case DATETIME -> {
                if      (value instanceof LocalDateTime ldt)    cell.setCellValue(ldt);
                else if (value instanceof java.sql.Timestamp ts) cell.setCellValue(ts.toLocalDateTime());
                else if (value instanceof java.util.Date d)     cell.setCellValue(d);
                else                                            cell.setCellValue(value.toString());
            }
            default -> cell.setCellValue(value.toString()); // STRING, ALPHANUMERIC
        }
    }

    private void tryDouble(Cell cell, String v) {
        try   { cell.setCellValue(Double.parseDouble(v)); }
        catch (NumberFormatException e) { cell.setCellValue(v); }
    }

    private void tryLong(Cell cell, String v) {
        try   { cell.setCellValue(Long.parseLong(v)); }
        catch (NumberFormatException e) { cell.setCellValue(v); }
    }

    private void tryDoubleDiv100(Cell cell, String v) {
        try   { cell.setCellValue(Double.parseDouble(v) / 100.0); }
        catch (NumberFormatException e) { cell.setCellValue(v); }
    }

    // -----------------------------------------------------------------------
    // Column width estimator for auto-sized columns (widthChars == 0)
    // -----------------------------------------------------------------------

    private int estimateWidth(ExcelColumnConfig col) {
        int headerChars = col.getHeader() != null ? col.getHeader().length() : 10;
        int typeChars = switch (col.getDataType()) {
            case CURRENCY    -> 16;
            case NUMERIC     -> 14;
            case INTEGER     -> 12;
            case DATE        -> 12;
            case DATETIME    -> 22;
            case PERCENTAGE  -> 10;
            case STRING, ALPHANUMERIC -> 25;
        };
        // +4 padding; * 256 = POI column width unit (1/256 of a character)
        return (Math.max(headerChars, typeChars) + 4) * 256;
    }
}
