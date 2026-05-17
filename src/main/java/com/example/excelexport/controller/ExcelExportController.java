package com.example.excelexport.controller;

import com.example.excelexport.controller.dto.ExcelExportRequest;
import com.example.excelexport.model.*;
import com.example.excelexport.service.DatabaseExcelDataProvider;
import com.example.excelexport.service.ExcelExportService;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.util.List;

/**
 * REST endpoints demonstrating the Excel export utility.
 *
 * Two patterns are shown:
 *
 * 1. Generic endpoint (POST /api/export/excel) — caller supplies SQL + columns at runtime.
 *    Useful for admin/reporting tools. Protect with role-based security; never expose publicly.
 *
 * 2. Screen-specific endpoint (GET /api/export/employees) — SQL and columns are hardcoded
 *    in the service layer. This is the recommended pattern for application screens.
 *
 * Both use {@link StreamingResponseBody} so Spring writes directly to the response output
 * stream on a separate thread, without buffering the entire file in heap memory.
 */
@Slf4j
@RestController
@RequestMapping("/api/export")
@RequiredArgsConstructor
public class ExcelExportController {

    private static final MediaType XLSX =
            MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

    private final ExcelExportService        excelExportService;
    private final DatabaseExcelDataProvider databaseExcelDataProvider;

    // -----------------------------------------------------------------------
    // Pattern 1 — generic, config-driven export
    // -----------------------------------------------------------------------

    /**
     * POST /api/export/excel
     *
     * Accepts a full export specification in the request body.
     * Secure this endpoint — it executes caller-supplied SQL against your database.
     */
    @PostMapping("/excel")
    public ResponseEntity<StreamingResponseBody> exportExcel(@RequestBody ExcelExportRequest req) {

        ExcelExportConfig config = ExcelExportConfig.builder()
                .sheetName(req.getSheetName())
                .fileName(req.getFileName())
                .columns(req.getColumns())
                .freezeColumnCount(req.getFreezeColumnCount())
                .freezeRowCount(req.getFreezeRowCount() > 0 ? req.getFreezeRowCount() : 1)
                .headerBgRgb(req.getHeaderBgRgb())
                .build();

        ExcelDataProvider data = databaseExcelDataProvider.fromQuery(req.getSql(), req.getParams());

        String fileName = StringUtils.hasText(req.getFileName()) ? req.getFileName() : "export.xlsx";
        return streamResponse(config, data, fileName);
    }

    // -----------------------------------------------------------------------
    // Pattern 2 — screen-specific export (copy this pattern per screen)
    // -----------------------------------------------------------------------

    /**
     * GET /api/export/employees?department=Engineering
     *
     * Demonstrates how any screen in the application wires up its own export.
     * Only the columns list, SQL, and config differ between screens — the
     * ExcelExportService and DatabaseExcelDataProvider are shared infrastructure.
     */
    @GetMapping("/employees")
    public ResponseEntity<StreamingResponseBody> exportEmployees(
            @RequestParam(required = false) String department) {

        // --- 1. Define columns ---
        List<ExcelColumnConfig> columns = List.of(
                col("Employee ID",       "employee_id",          ColumnDataType.INTEGER,     12),
                col("Full Name",         "full_name",            ColumnDataType.STRING,      30),
                col("Department",        "department_name",      ColumnDataType.STRING,      20),
                col("Job Title",         "job_title",            ColumnDataType.ALPHANUMERIC, 25),
                col("Salary",            "salary",               ColumnDataType.CURRENCY,    15),
                col("Hire Date",         "hire_date",            ColumnDataType.DATE,        12),
                col("Performance (%)",   "performance_rating",   ColumnDataType.PERCENTAGE,  14)
        );

        // --- 2. Build config ---
        ExcelExportConfig config = ExcelExportConfig.builder()
                .sheetName("Employees")
                .fileName("employees_export.xlsx")
                .columns(columns)
                .freezeColumnCount(2)   // Employee ID + Full Name always visible while scrolling right
                .freezeRowCount(1)      // Header row always visible while scrolling down
                .headerBgRgb(new int[]{31, 73, 125}) // optional — matches DEFAULT_HEADER_BG
                .build();

        // --- 3. Build parameterised query ---
        String baseSql = """
                SELECT e.employee_id,
                       e.full_name,
                       d.department_name,
                       e.job_title,
                       e.salary,
                       e.hire_date,
                       e.performance_rating
                  FROM employees   e
                  JOIN departments d ON d.dept_id = e.dept_id
                """;

        String sql;
        Object[] params;
        if (department != null) {
            sql    = baseSql + " WHERE d.department_name = ? ORDER BY e.full_name";
            params = new Object[]{department};
        } else {
            sql    = baseSql + " ORDER BY e.full_name";
            params = new Object[0];
        }

        // --- 4. Stream ---
        ExcelDataProvider data = databaseExcelDataProvider.fromQuery(sql, params);
        return streamResponse(config, data, "employees_export.xlsx");
    }

    // -----------------------------------------------------------------------
    // Shared response builder
    // -----------------------------------------------------------------------

    private ResponseEntity<StreamingResponseBody> streamResponse(ExcelExportConfig config,
                                                                   ExcelDataProvider data,
                                                                   String fileName) {
        StreamingResponseBody body = out -> {
            try {
                excelExportService.exportToStream(config, data, out);
            } catch (Exception e) {
                log.error("Excel export failed for '{}'", fileName, e);
                throw new RuntimeException("Excel export failed", e);
            }
        };

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + fileName + "\"")
                .contentType(XLSX)
                .body(body);
    }

    // -----------------------------------------------------------------------
    // Column builder shortcut
    // -----------------------------------------------------------------------

    private ExcelColumnConfig col(String header, String field, ColumnDataType type, int width) {
        return ExcelColumnConfig.builder()
                .header(header)
                .fieldName(field)
                .dataType(type)
                .widthChars(width)
                .build();
    }
}
