package com.example.excelexport.model;

/**
 * Strategy that feeds rows to the Excel writer.
 *
 * Implementations must call {@code rowWriter.writeRow(map)} once per record.
 * The map key is the lower-cased SQL column label; value is the raw JDBC object.
 *
 * Two built-in implementations are available:
 * <ul>
 *   <li>{@code DatabaseExcelDataProvider.fromQuery(sql, params)} — streams directly from PostgreSQL</li>
 *   <li>{@code ExcelDataProvider.ofList(list)} — wraps an in-memory list (small datasets)</li>
 * </ul>
 */
@FunctionalInterface
public interface ExcelDataProvider {
    void fetchData(ExcelRowWriter rowWriter) throws Exception;

    /** Convenience factory for small in-memory datasets. */
    static ExcelDataProvider ofList(Iterable<java.util.Map<String, Object>> rows) {
        return rowWriter -> rows.forEach(rowWriter::writeRow);
    }
}
