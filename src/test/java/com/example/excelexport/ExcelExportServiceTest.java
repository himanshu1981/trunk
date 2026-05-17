package com.example.excelexport;

import com.example.excelexport.model.*;
import com.example.excelexport.service.ExcelExportService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.time.LocalDate;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class ExcelExportServiceTest {

    private final ExcelExportService service = new ExcelExportService();

    @Test
    void headerRowContainsAllColumnNames() throws Exception {
        ExcelExportConfig config = buildConfig(List.of(
                col("Name",   "name",   ColumnDataType.STRING),
                col("Amount", "amount", ColumnDataType.CURRENCY),
                col("Date",   "dt",     ColumnDataType.DATE)
        ));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        service.exportToStream(config, ExcelDataProvider.ofList(List.of()), out);

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Row header = wb.getSheetAt(0).getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("Name");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("Amount");
            assertThat(header.getCell(2).getStringCellValue()).isEqualTo("Date");
        }
    }

    @Test
    void dataRowsAreWrittenInOrder() throws Exception {
        List<Map<String, Object>> rows = List.of(
                Map.of("id", 1L, "name", "Alice", "salary", 80_000.0),
                Map.of("id", 2L, "name", "Bob",   "salary", 92_000.0)
        );

        ExcelExportConfig config = buildConfig(List.of(
                col("ID",     "id",     ColumnDataType.INTEGER),
                col("Name",   "name",   ColumnDataType.STRING),
                col("Salary", "salary", ColumnDataType.CURRENCY)
        ));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        service.exportToStream(config, ExcelDataProvider.ofList(rows), out);

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getLastRowNum()).isEqualTo(2); // row 0 = header, rows 1-2 = data
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("Alice");
            assertThat(sheet.getRow(2).getCell(1).getStringCellValue()).isEqualTo("Bob");
            assertThat(sheet.getRow(1).getCell(2).getNumericCellValue()).isEqualTo(80_000.0);
        }
    }

    @Test
    void nullValuesProduceBlankCells() throws Exception {
        List<Map<String, Object>> rows = List.of(Map.of("id", 1L));

        ExcelExportConfig config = buildConfig(List.of(
                col("ID",   "id",   ColumnDataType.INTEGER),
                col("Name", "name", ColumnDataType.STRING)   // "name" absent from map → null
        ));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        service.exportToStream(config, ExcelDataProvider.ofList(rows), out);

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Cell nameCell = wb.getSheetAt(0).getRow(1).getCell(1);
            assertThat(nameCell.getCellType()).isEqualTo(CellType.BLANK);
        }
    }

    @Test
    void localDateIsWrittenAsDateCell() throws Exception {
        LocalDate date = LocalDate.of(2024, 6, 15);
        List<Map<String, Object>> rows = List.of(Map.of("dt", date));

        ExcelExportConfig config = buildConfig(List.of(col("Date", "dt", ColumnDataType.DATE)));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        service.exportToStream(config, ExcelDataProvider.ofList(rows), out);

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Cell cell = wb.getSheetAt(0).getRow(1).getCell(0);
            assertThat(cell.getCellType()).isEqualTo(CellType.NUMERIC);
            assertThat(DateUtil.isCellDateFormatted(cell)).isTrue();
        }
    }

    @Test
    void percentageIsDividedBy100() throws Exception {
        List<Map<String, Object>> rows = List.of(Map.of("pct", 75.0));

        ExcelExportConfig config = buildConfig(List.of(col("Score", "pct", ColumnDataType.PERCENTAGE)));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        service.exportToStream(config, ExcelDataProvider.ofList(rows), out);

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            double stored = wb.getSheetAt(0).getRow(1).getCell(0).getNumericCellValue();
            assertThat(stored).isEqualTo(0.75); // 75 / 100
        }
    }

    // -----------------------------------------------------------------------

    private ExcelExportConfig buildConfig(List<ExcelColumnConfig> columns) {
        return ExcelExportConfig.builder()
                .sheetName("Test")
                .columns(columns)
                .freezeColumnCount(1)
                .build();
    }

    private ExcelColumnConfig col(String header, String field, ColumnDataType type) {
        return ExcelColumnConfig.builder()
                .header(header).fieldName(field).dataType(type).build();
    }
}
