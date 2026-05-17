package com.example.excelexport.controller.dto;

import com.example.excelexport.model.ExcelColumnConfig;
import lombok.Data;

import java.util.List;

/**
 * Request body for the generic POST /api/export/excel endpoint.
 *
 * JSON example:
 * <pre>
 * {
 *   "sheetName": "Orders",
 *   "fileName": "orders_2024.xlsx",
 *   "sql": "SELECT order_id, customer_name, total FROM orders WHERE year = ?",
 *   "params": [2024],
 *   "freezeColumnCount": 1,
 *   "headerBgRgb": [0, 112, 192],
 *   "columns": [
 *     { "header": "Order ID",      "fieldName": "order_id",       "dataType": "INTEGER"  },
 *     { "header": "Customer",      "fieldName": "customer_name",  "dataType": "STRING"   },
 *     { "header": "Total",         "fieldName": "total",          "dataType": "CURRENCY" }
 *   ]
 * }
 * </pre>
 */
@Data
public class ExcelExportRequest {

    /** Parameterised SQL — values supplied via {@code params}. */
    private String sql;

    /** Positional bind values matching each {@code ?} in {@code sql}. */
    private Object[] params;

    /** Column definitions — order determines Excel column order. */
    private List<ExcelColumnConfig> columns;

    private String sheetName;
    private String fileName;

    /** Number of columns to freeze from the left (0 = none). */
    private int freezeColumnCount;

    /** Number of rows to freeze from the top (0 defaults to 1 = header). */
    private int freezeRowCount;

    /** RGB header background. Null = default dark blue. Example: [0, 176, 240] */
    private int[] headerBgRgb;
}
