package com.example.excelexport.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * Top-level configuration for one export operation.
 *
 * Minimal usage (all optional fields have sensible defaults):
 * <pre>
 *   ExcelExportConfig config = ExcelExportConfig.builder()
 *       .sheetName("Sales Report")
 *       .fileName("sales_2024.xlsx")
 *       .columns(columnList)
 *       .freezeColumnCount(2)   // freeze first 2 columns
 *       .build();
 * </pre>
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExcelExportConfig {

    /** Worksheet tab name. Defaults to "Sheet1". */
    private String sheetName;

    /** Suggested file name sent in the Content-Disposition header. */
    private String fileName;

    /** Ordered list of column definitions. */
    private List<ExcelColumnConfig> columns;

    /**
     * Number of columns to freeze from the left edge (freeze pane).
     * 0 = no column freeze.
     */
    @Builder.Default
    private int freezeColumnCount = 0;

    /**
     * Number of rows to freeze from the top.
     * Defaults to 1 so the header row is always visible.
     */
    @Builder.Default
    private int freezeRowCount = 1;

    /**
     * RGB values [R, G, B] for the header background, each 0-255.
     * Null = default corporate dark blue {31, 73, 125}.
     *
     * Example — teal header: new int[]{0, 128, 128}
     */
    private int[] headerBgRgb;

    /**
     * SXSSF in-memory sliding window size (rows kept in RAM at once).
     * Lower = less heap usage; higher = faster temp-file I/O.
     * Defaults to 1 000. Suitable for 1 M+ row exports.
     */
    @Builder.Default
    private int rowAccessWindowSize = 1_000;

    /**
     * Maximum data rows per sheet before a new sheet is created automatically.
     * Set this when the product of (rows × columns) would push uncompressed sheet XML
     * beyond ~1.8 GB — the safe ceiling below POI's ZIP integer-overflow threshold.
     *
     * Rule of thumb: maxRowsPerSheet = floor(1_800_000_000 / (columns × 45))
     *   e.g. 340 cols → ~117 000 rows per sheet
     *        60  cols → ~666 000 rows per sheet
     *
     * 0 (default) = unlimited / single sheet.
     * Each continuation sheet is named "SheetName (2)", "SheetName (3)", …
     */
    @Builder.Default
    private int maxRowsPerSheet = 0;
}
