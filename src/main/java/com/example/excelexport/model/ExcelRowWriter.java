package com.example.excelexport.model;

import java.util.Map;

/**
 * Callback invoked once per data row by an {@link ExcelDataProvider}.
 * Keys in the map must match {@link ExcelColumnConfig#getFieldName()}.
 */
@FunctionalInterface
public interface ExcelRowWriter {
    void writeRow(Map<String, Object> rowData);
}
