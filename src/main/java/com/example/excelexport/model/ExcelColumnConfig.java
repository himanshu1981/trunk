package com.example.excelexport.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * Describes a single column in the exported spreadsheet.
 *
 * Usage example:
 * <pre>
 *   ExcelColumnConfig.builder()
 *       .header("Invoice Amount")
 *       .fieldName("invoice_amount")   // must match the Map key returned by the query
 *       .dataType(ColumnDataType.CURRENCY)
 *       .widthChars(16)
 *       .build()
 * </pre>
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExcelColumnConfig {

    /** Text shown in the header row. */
    private String header;

    /**
     * Key used to look up the value in each data row's Map.
     * Typically the lower-cased SQL column label (e.g. "invoice_amount").
     */
    private String fieldName;

    /** Controls cell format mask and horizontal alignment. */
    private ColumnDataType dataType;

    /**
     * Explicit column width in character units.
     * Set to 0 (default) to let the utility auto-estimate based on data type and header length.
     */
    @Builder.Default
    private int widthChars = 0;

    /**
     * Optional Excel number format mask that overrides the dataType default.
     * Example: "dd/MM/yyyy", "#,##0.000", "€#,##0.00"
     */
    private String numberFormat;
}
