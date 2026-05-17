package com.example.excelexport.model;

/**
 * Drives cell formatting, number format masks, and alignment.
 *
 * International alignment rules applied automatically:
 *   Numeric / Integer / Currency / Percentage / Date / Datetime → RIGHT
 *   String / Alphanumeric                                        → LEFT
 */
public enum ColumnDataType {
    STRING,       // free text
    ALPHANUMERIC, // codes, IDs that mix digits and letters but must not be parsed as numbers
    INTEGER,      // whole numbers          — format: #,##0
    NUMERIC,      // decimals               — format: #,##0.##
    CURRENCY,     // monetary amounts       — format: $#,##0.00
    PERCENTAGE,   // ratio stored as 0-100  — format: 0.00%  (value divided by 100 before write)
    DATE,         // java.sql.Date / LocalDate             — format: yyyy-mm-dd
    DATETIME      // java.sql.Timestamp / LocalDateTime    — format: yyyy-mm-dd hh:mm:ss
}
