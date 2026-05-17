package com.example.excelexport.service;

import com.example.excelexport.model.ExcelDataProvider;
import lombok.extern.slf4j.Slf4j;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;
import org.springframework.transaction.PlatformTransactionManager;
import org.springframework.transaction.TransactionDefinition;
import org.springframework.transaction.support.TransactionTemplate;

import java.sql.ResultSetMetaData;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Builds {@link ExcelDataProvider} instances that stream rows directly from PostgreSQL
 * using a server-side cursor, keeping memory usage flat regardless of result set size.
 *
 * <p><b>How PostgreSQL cursor streaming works:</b><br>
 * Setting {@code fetchSize > 0} on a {@link java.sql.PreparedStatement} inside an active
 * transaction instructs the PostgreSQL JDBC driver to fetch rows in batches (e.g. 1 000 at
 * a time) rather than buffering the entire result set in memory.  A read-only transaction is
 * opened automatically; no caller setup is required.
 *
 * <p><b>Usage from a controller or service:</b>
 * <pre>
 *   ExcelDataProvider data = databaseExcelDataProvider.fromQuery(
 *       "SELECT id, name, amount FROM invoices WHERE status = ?", "PAID");
 *
 *   excelExportService.exportToStream(config, data, response.getOutputStream());
 * </pre>
 */
@Slf4j
@Service
public class DatabaseExcelDataProvider {

    private static final int FETCH_SIZE = 1_000;

    private final JdbcTemplate      jdbcTemplate;
    private final TransactionTemplate readOnlyTx;

    public DatabaseExcelDataProvider(JdbcTemplate jdbcTemplate,
                                     PlatformTransactionManager txManager) {
        this.jdbcTemplate = jdbcTemplate;
        this.readOnlyTx   = new TransactionTemplate(txManager);
        this.readOnlyTx.setReadOnly(true);
        this.readOnlyTx.setPropagationBehavior(TransactionDefinition.PROPAGATION_REQUIRED);
    }

    /**
     * Creates a streaming data provider for the given SQL query.
     *
     * @param sql    Parameterised SQL — use {@code ?} placeholders
     * @param params Values for each {@code ?} placeholder, in order
     * @return an {@link ExcelDataProvider} ready to stream rows to the Excel writer
     */
    public ExcelDataProvider fromQuery(String sql, Object... params) {
        return rowWriter -> readOnlyTx.execute(status -> {
            log.debug("Streaming query for Excel export: {}", sql);

            jdbcTemplate.query(
                con -> {
                    // TYPE_FORWARD_ONLY + CONCUR_READ_ONLY is mandatory for PostgreSQL cursors
                    var ps = con.prepareStatement(
                            sql,
                            java.sql.ResultSet.TYPE_FORWARD_ONLY,
                            java.sql.ResultSet.CONCUR_READ_ONLY
                    );
                    ps.setFetchSize(FETCH_SIZE); // triggers server-side cursor in PostgreSQL
                    if (params != null) {
                        for (int i = 0; i < params.length; i++) {
                            ps.setObject(i + 1, params[i]);
                        }
                    }
                    return ps;
                },
                rs -> {
                    ResultSetMetaData meta = rs.getMetaData();
                    int colCount = meta.getColumnCount();

                    // LinkedHashMap preserves SQL column order (useful for debugging)
                    Map<String, Object> row = new LinkedHashMap<>(colCount);
                    for (int i = 1; i <= colCount; i++) {
                        // Lower-case the label so ExcelColumnConfig.fieldName can use either case
                        row.put(meta.getColumnLabel(i).toLowerCase(), rs.getObject(i));
                    }
                    rowWriter.writeRow(row);
                }
            );
            return null;
        });
    }
}
