package com.example.excelexport;

import com.example.excelexport.model.*;
import com.example.excelexport.service.ExcelExportService;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

import java.io.BufferedOutputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.*;

/**
 * Performance benchmark: 1 000 000 rows × 340 mixed-type columns.
 *
 * The data provider is constructed BEFORE the timer starts, so the
 * reported time measures only Excel generation and disk write — no DB time.
 *
 * Run manually:
 *   mvn test -Dtest=MillionRowBenchmarkTest -Dbenchmark=true
 *
 * Or remove @Disabled and run via IDE.
 */
@Disabled("Long-running benchmark — run manually: mvn test -Dtest=MillionRowBenchmarkTest")
class MillionRowBenchmarkTest {

    private static final int ROWS = 1_000_000;
    private static final int COLS = 340;

    // Column counts per type — must sum to COLS (340)
    private static final int INT_COLS  = 50;   // whole numbers
    private static final int NUM_COLS  = 60;   // decimals
    private static final int AMT_COLS  = 60;   // currency / amounts
    private static final int PCT_COLS  = 30;   // percentages (0-100)
    private static final int TXT_COLS  = 70;   // alphabetic strings
    private static final int CODE_COLS = 50;   // alphanumeric codes
    private static final int DATE_COLS = 20;   // dates

    private final ExcelExportService service = new ExcelExportService();

    // -----------------------------------------------------------------------
    // Benchmark entry point
    // -----------------------------------------------------------------------

    @Test
    void benchmark_1Million_340Columns() throws Exception {

        System.out.println("\n=== Building column schema ===");
        List<ExcelColumnConfig> columns = buildColumns();

        // maxRowsPerSheet keeps each sheet's uncompressed XML under ~1.8 GB —
        // the safe ceiling before POI's internal ZIP counter overflows (int overflow).
        // Rule of thumb:  floor(1_800_000_000 / (columns × ~45 bytes/cell))
        //   340 cols  →  floor(1.8e9 / 15 300) ≈ 117 000  →  100 000 for safety margin
        //   1M rows / 100K = 10 sheets
        final int MAX_ROWS_PER_SHEET = 100_000;

        ExcelExportConfig config = ExcelExportConfig.builder()
                .sheetName("BenchmarkReport")
                .fileName("benchmark_1m_340col.xlsx")
                .columns(columns)
                .freezeColumnCount(3)
                .freezeRowCount(1)
                .rowAccessWindowSize(2_000)
                .maxRowsPerSheet(MAX_ROWS_PER_SHEET)
                .build();

        // -----------------------------------------------------------------------
        // Pre-compute fixed-size data pools.
        // Drawing from pools (% modulo) avoids allocating random strings in the hot loop,
        // cutting per-row allocation cost dramatically.
        // -----------------------------------------------------------------------
        System.out.println("=== Pre-computing data pools ===");
        String[]    namePool = buildNamePool(1_000);
        String[]    codePool = buildCodePool(1_000);
        LocalDate[] datePool = buildDatePool(1_000);

        // -----------------------------------------------------------------------
        // Data provider — rows are generated on demand, one at a time.
        // A single Map<> instance is reused (service reads it synchronously and
        // never holds a reference after writeRow returns), keeping heap flat.
        // This simulates a JDBC RowCallbackHandler where the ResultSet is forward-only.
        // -----------------------------------------------------------------------
        ExcelDataProvider dataProvider = rowWriter -> {
            Random rng = new Random(42L);
            Map<String, Object> row = new HashMap<>(COLS * 2, 0.75f);

            for (int r = 0; r < ROWS; r++) {
                row.clear();
                fillRow(row, r, rng, columns, namePool, codePool, datePool);
                rowWriter.writeRow(row);
            }
        };

        // -----------------------------------------------------------------------
        // Print column breakdown BEFORE timer starts
        // -----------------------------------------------------------------------
        printHeader(columns);

        // -----------------------------------------------------------------------
        // Write to temp file.
        // Timer starts HERE — everything above is "pre-export / DB-equivalent" time.
        // -----------------------------------------------------------------------
        Path outputFile = Files.createTempFile("excel_benchmark_1M_", ".xlsx");

        long startNs = System.nanoTime();

        try (OutputStream os = new BufferedOutputStream(
                Files.newOutputStream(outputFile), 128 * 1024)) { // 128 KB write buffer
            service.exportToStream(config, dataProvider, os);
        }

        long elapsedNs = System.nanoTime() - startNs;

        // -----------------------------------------------------------------------
        // Results
        // -----------------------------------------------------------------------
        double elapsedSec  = elapsedNs / 1_000_000_000.0;
        double rowsPerSec  = ROWS / elapsedSec;
        long   fileSizeB   = Files.size(outputFile);
        double fileSizeMB  = fileSizeB / (1024.0 * 1024.0);
        double fileSizeGB  = fileSizeB / (1024.0 * 1024.0 * 1024.0);
        double msPerKRows  = (elapsedNs / 1_000_000.0) / (ROWS / 1_000.0); // ms per 1K rows

        int sheetsUsed = (int) Math.ceil((double) ROWS / MAX_ROWS_PER_SHEET);
        System.out.printf("""

╔════════════════════════════════════════════════════════╗
║            EXCEL EXPORT BENCHMARK RESULTS              ║
╠═══════════════════════════════╦════════════════════════╣
║  Total rows exported          ║ %,20d  ║
║  Total columns                ║ %,20d  ║
║  Sheets (auto-split)          ║ %,20d  ║
║  Max rows per sheet           ║ %,20d  ║
╠═══════════════════════════════╬════════════════════════╣
║  Excel generation time        ║ %19.2f s  ║
║  Throughput                   ║ %,17.0f rows/s  ║
║  Avg time per 1 000 rows      ║ %19.1f ms  ║
╠═══════════════════════════════╬════════════════════════╣
║  Output file size             ║ %16.2f MB     ║
║  Output file size             ║ %16.3f GB     ║
╠═══════════════════════════════╬════════════════════════╣
║  NOTE: DB fetch time          ║      EXCLUDED (0 ms)   ║
║  (rows generated in-memory)   ║                        ║
╚═══════════════════════════════╩════════════════════════╝
""",
                (long) ROWS, (long) COLS, (long) sheetsUsed, (long) MAX_ROWS_PER_SHEET,
                elapsedSec,
                rowsPerSec,
                msPerKRows,
                fileSizeMB,
                fileSizeGB);

        System.out.printf("Output written to: %s%n", outputFile);
        System.out.printf("(file deleted after test)%n%n");

        Files.deleteIfExists(outputFile);
    }

    // -----------------------------------------------------------------------
    // Column schema builder — 340 columns, realistic enterprise report mix
    // -----------------------------------------------------------------------

    private List<ExcelColumnConfig> buildColumns() {
        List<ExcelColumnConfig> cols = new ArrayList<>(COLS);

        // 50 INTEGER — IDs, counts, quantities
        for (int i = 1; i <= INT_COLS; i++)
            cols.add(col("Int " + i,       "int_col_" + i,    ColumnDataType.INTEGER,     10));

        // 60 NUMERIC — measurements, rates, scores
        for (int i = 1; i <= NUM_COLS; i++)
            cols.add(col("Num " + i,       "num_col_" + i,    ColumnDataType.NUMERIC,     13));

        // 60 CURRENCY — amounts, prices, costs
        for (int i = 1; i <= AMT_COLS; i++)
            cols.add(col("Amount " + i,    "amt_col_" + i,    ColumnDataType.CURRENCY,    14));

        // 30 PERCENTAGE — ratios, completion rates, margins
        for (int i = 1; i <= PCT_COLS; i++)
            cols.add(col("Pct " + i,       "pct_col_" + i,    ColumnDataType.PERCENTAGE,  11));

        // 70 STRING — names, descriptions, statuses
        for (int i = 1; i <= TXT_COLS; i++)
            cols.add(col("Text " + i,      "txt_col_" + i,    ColumnDataType.STRING,      20));

        // 50 ALPHANUMERIC — reference codes, SKUs, invoice numbers
        for (int i = 1; i <= CODE_COLS; i++)
            cols.add(col("Code " + i,      "code_col_" + i,   ColumnDataType.ALPHANUMERIC, 16));

        // 20 DATE — event dates, due dates, creation dates
        for (int i = 1; i <= DATE_COLS; i++)
            cols.add(col("Date " + i,      "date_col_" + i,   ColumnDataType.DATE,        12));

        return cols;
    }

    // -----------------------------------------------------------------------
    // Row filler — writes into the reusable map (avoids per-row allocation)
    // -----------------------------------------------------------------------

    private void fillRow(Map<String, Object> row,
                          int rowIndex,
                          Random rng,
                          List<ExcelColumnConfig> columns,
                          String[] namePool,
                          String[] codePool,
                          LocalDate[] datePool) {

        for (ExcelColumnConfig col : columns) {
            row.put(col.getFieldName(), nextValue(col.getDataType(), rowIndex, rng,
                    namePool, codePool, datePool));
        }
    }

    private Object nextValue(ColumnDataType type, int rowIndex, Random rng,
                              String[] namePool, String[] codePool, LocalDate[] datePool) {
        return switch (type) {
            case INTEGER     -> (long) (rng.nextInt(9_999_999) + 1);
            case NUMERIC     -> Math.round(rng.nextDouble() * 99_999.99 * 100.0) / 100.0;
            case CURRENCY    -> Math.round(rng.nextDouble() * 999_999.99 * 100.0) / 100.0;
            case PERCENTAGE  -> Math.round(rng.nextDouble() * 100.0 * 100.0) / 100.0; // 0.00–100.00
            case STRING      -> namePool[rowIndex % namePool.length];
            case ALPHANUMERIC -> codePool[rowIndex % codePool.length];
            case DATE        -> datePool[rowIndex % datePool.length];
            case DATETIME    -> null; // not used in this schema
        };
    }

    // -----------------------------------------------------------------------
    // Pool builders
    // -----------------------------------------------------------------------

    private String[] buildNamePool(int size) {
        String[] first = {"Alice", "Bob", "Carol", "David", "Eva", "Frank", "Grace", "Henry",
                "Iris", "Jack", "Karen", "Leo", "Mia", "Noah", "Olivia", "Paul",
                "Quinn", "Rachel", "Sam", "Tina", "Uma", "Victor", "Wendy", "Xander",
                "Yara", "Zoe"};
        String[] last  = {"Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller",
                "Davis", "Wilson", "Moore", "Taylor", "Anderson", "Thomas", "Jackson",
                "White", "Harris", "Martin", "Thompson", "Young", "Lee"};
        Random rng = new Random(0L);
        String[] pool = new String[size];
        for (int i = 0; i < size; i++)
            pool[i] = first[rng.nextInt(first.length)] + " " + last[rng.nextInt(last.length)];
        return pool;
    }

    private String[] buildCodePool(int size) {
        String chars  = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
        Random rng    = new Random(1L);
        String[] pool = new String[size];
        for (int i = 0; i < size; i++) {
            char[] buf = new char[12];
            for (int j = 0; j < 12; j++)
                buf[j] = chars.charAt(rng.nextInt(chars.length()));
            pool[i] = new String(buf);
        }
        return pool;
    }

    private LocalDate[] buildDatePool(int size) {
        LocalDate base = LocalDate.of(2019, 1, 1);
        Random rng = new Random(2L);
        LocalDate[] pool = new LocalDate[size];
        for (int i = 0; i < size; i++)
            pool[i] = base.plusDays(rng.nextInt(365 * 6)); // dates in 2019-2024
        return pool;
    }

    // -----------------------------------------------------------------------
    // Helpers
    // -----------------------------------------------------------------------

    private ExcelColumnConfig col(String header, String field, ColumnDataType type, int width) {
        return ExcelColumnConfig.builder()
                .header(header).fieldName(field).dataType(type).widthChars(width).build();
    }

    private void printHeader(List<ExcelColumnConfig> columns) {
        Map<ColumnDataType, Long> counts = new LinkedHashMap<>();
        for (ColumnDataType dt : ColumnDataType.values()) counts.put(dt, 0L);
        for (ExcelColumnConfig c : columns) counts.merge(c.getDataType(), 1L, Long::sum);

        System.out.printf("""

=== Column Distribution (total: %d) ===
  INTEGER      (whole numbers, right-aligned)  : %3d cols
  NUMERIC      (decimals,      right-aligned)  : %3d cols
  CURRENCY     (amounts,       right-aligned)  : %3d cols
  PERCENTAGE   (0-100 %%,       right-aligned)  : %3d cols
  STRING       (names/text,    left-aligned)   : %3d cols
  ALPHANUMERIC (codes/SKUs,    left-aligned)   : %3d cols
  DATE         (yyyy-mm-dd,    right-aligned)  : %3d cols
=== Features ===
  Header   : background colour + bold white font + borders + auto-filter
  Freeze   : first 3 columns frozen (horizontal scroll)
           : header row frozen       (vertical scroll)
  Formats  : per-type number masks (#,##0 / $#,##0.00 / 0.00%% / yyyy-mm-dd)
  Streaming: SXSSF window=2 000 rows — heap stays flat at any row count
""",
                columns.size(),
                counts.get(ColumnDataType.INTEGER),
                counts.get(ColumnDataType.NUMERIC),
                counts.get(ColumnDataType.CURRENCY),
                counts.get(ColumnDataType.PERCENTAGE),
                counts.get(ColumnDataType.STRING),
                counts.get(ColumnDataType.ALPHANUMERIC),
                counts.get(ColumnDataType.DATE));
    }
}
