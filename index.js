'use strict';

const express = require('express');
const ExcelJS = require('exceljs');
const pino = require('pino');

const logger = pino({
  level: process.env.LOG_LEVEL || 'info',
  ...(process.env.NODE_ENV !== 'production' && { transport: { target: 'pino-pretty' } }),
});

const app = express();
const PORT = process.env.PORT || 3000;

// ─── Constants ────────────────────────────────────────────────────────────────

const PRIORITY_HEADERS = ['Record_ID', 'Field_Name', 'Old_Value', 'New_Value'];
const MAX_COL_WIDTH = 50;
const MIN_COL_WIDTH = 10;
const WRAP_TEXT_THRESHOLD = 50; // characters; cells longer than this get wrapText enabled

// ─── Middleware ───────────────────────────────────────────────────────────────

app.use(express.json({ limit: '10mb' }));

// ─── Routes ───────────────────────────────────────────────────────────────────

app.get('/health', (_req, res) => {
  res.json({ status: 'ok', ts: new Date().toISOString() });
});

app.post('/generate-excel', async (req, res) => {
  const start = Date.now();
  const { rows, fileName = 'salesforce_report.xlsx', sheetName = 'Report' } = req.body || {};

  // ── Validation ──────────────────────────────────────────────────────────────

  if (!rows || !Array.isArray(rows)) {
    logger.warn('Invalid request: rows missing or not an array');
    return res.status(400).json({ error: 'rows must be a non-empty array' });
  }

  if (rows.length === 0) {
    logger.warn('Invalid request: rows is empty');
    return res.status(400).json({ error: 'rows must not be empty' });
  }

  if (!rows.every((r) => r !== null && typeof r === 'object' && !Array.isArray(r))) {
    logger.warn('Invalid request: one or more rows is not a plain object');
    return res.status(400).json({ error: 'each row must be a plain object' });
  }

  // ── Header ordering ─────────────────────────────────────────────────────────
  // Build a union of all keys across all rows.
  // Priority headers come first (in spec order), then any additional keys in
  // the order they first appear.

  const seenKeys = new Set(PRIORITY_HEADERS);
  const extraKeys = [];

  rows.forEach((row) => {
    Object.keys(row).forEach((k) => {
      if (!seenKeys.has(k)) {
        seenKeys.add(k);
        extraKeys.push(k);
      }
    });
  });

  const allRowKeys = new Set(rows.flatMap((r) => Object.keys(r)));
  const headers = [
    ...PRIORITY_HEADERS.filter((h) => allRowKeys.has(h)),
    ...extraKeys,
  ];

  if (headers.length === 0) {
    logger.warn('Invalid request: rows have no fields');
    return res.status(400).json({ error: 'rows must contain at least one field' });
  }

  // ── sheetName sanitization ──────────────────────────────────────────────────
  // Excel sheet names cannot contain: [ ] : * ? / \
  // and must be <= 31 characters.

  const rawName = typeof sheetName === 'string' ? sheetName : 'Report';
  const safeName = rawName.replace(/[[\]:*?/\\]/g, '').slice(0, 31).trim() || 'Report';

  if (safeName !== rawName) {
    logger.warn({ rawName, safeName }, 'sheetName was sanitized');
  }

  // ── safeFileName for Content-Disposition ────────────────────────────────────

  const rawFileName = typeof fileName === 'string' && fileName.trim()
    ? fileName.trim()
    : 'salesforce_report.xlsx';
  const safeFileName = rawFileName.replace(/[^\w.\-]/g, '_');

  logger.info({ rowCount: rows.length, colCount: headers.length, fileName: safeFileName, sheetName: safeName }, 'Generating Excel workbook');

  // ── Workbook generation ─────────────────────────────────────────────────────

  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(safeName);

    // Define columns — ExcelJS writes the header row automatically when
    // `header` is provided. Data rows begin at row 2.
    worksheet.columns = headers.map((h) => ({
      header: h,
      key: h,
      width: MIN_COL_WIDTH,
    }));

    // Bold header row (row 1)
    worksheet.getRow(1).font = { bold: true };

    // Freeze the top row so it stays visible while scrolling
    worksheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 1, topLeftCell: 'A2' }];

    // Autofilter across all header columns
    worksheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: headers.length },
    };

    // Add data rows
    rows.forEach((row) => {
      const values = headers.map((h) => {
        const v = row[h];
        return v !== undefined && v !== null ? v : '';
      });

      const excelRow = worksheet.addRow(values);

      // Enable wrap text for cells with longer content
      excelRow.eachCell({ includeEmpty: true }, (cell) => {
        const strVal = cell.value !== null && cell.value !== undefined
          ? String(cell.value)
          : '';
        if (strVal.length > WRAP_TEXT_THRESHOLD) {
          cell.alignment = { wrapText: true, vertical: 'top' };
        }
      });
    });

    // Auto-size columns based on the longest value in each column
    // (including the header text), capped at MAX_COL_WIDTH.
    headers.forEach((_header, colIndex) => {
      const column = worksheet.getColumn(colIndex + 1);
      let maxLen = 0;

      column.eachCell({ includeEmpty: false }, (cell) => {
        const len = cell.value !== null && cell.value !== undefined
          ? String(cell.value).length
          : 0;
        if (len > maxLen) maxLen = len;
      });

      column.width = Math.min(Math.max(maxLen + 2, MIN_COL_WIDTH), MAX_COL_WIDTH);
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const durationMs = Date.now() - start;

    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="${safeFileName}"`,
      'Content-Length': buffer.length,
      'X-Content-Type-Options': 'nosniff',
    });

    logger.info({ rowCount: rows.length, bufferBytes: buffer.length, durationMs }, 'Excel workbook generated successfully');

    res.end(buffer);
  } catch (err) {
    logger.error({ err: err.message, stack: err.stack, rowCount: rows.length }, 'Excel generation failed');
    res.status(500).json({ error: 'Failed to generate Excel file' });
  }
});

// ─── 404 ──────────────────────────────────────────────────────────────────────

app.use((_req, res) => {
  res.status(404).json({ error: 'Not found' });
});

// ─── Error handler ────────────────────────────────────────────────────────────

// eslint-disable-next-line no-unused-vars
app.use((err, _req, res, _next) => {
  logger.error({ err: err.message }, 'Unhandled error');
  res.status(err.status || 500).json({ error: err.message || 'Internal server error' });
});

// ─── Start ────────────────────────────────────────────────────────────────────

if (require.main === module) {
  app.listen(PORT, () => {
    logger.info({ port: PORT }, 'Excel formatter service started');
  });
}

module.exports = app;
