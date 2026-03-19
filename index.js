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
  const { rows, fileName = 'salesforce_report.xlsx', sheetName = 'Report', title, subtitle } = req.body || {};

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

    // ── Optional title/subtitle header block ─────────────────────────────────
    // If a title (or subtitle) is provided, insert rows above the data header.
    // headerOffset tracks how many rows were prepended so we can shift the
    // freeze pane and autofilter down accordingly.

    let headerOffset = 0;

    if (title || subtitle) {
      const colCount = headers.length;

      const headerFill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD9E1F2' }, // light blue-grey
      };

      if (title) {
        const titleRow = worksheet.addRow([String(title)]);
        worksheet.mergeCells(titleRow.number, 1, titleRow.number, colCount);
        titleRow.getCell(1).font = { bold: true, size: 13 };
        titleRow.getCell(1).alignment = { vertical: 'middle' };
        titleRow.getCell(1).fill = headerFill;
        titleRow.height = 20;
        headerOffset += 1;
      }

      if (subtitle) {
        const subtitleRow = worksheet.addRow([String(subtitle)]);
        worksheet.mergeCells(subtitleRow.number, 1, subtitleRow.number, colCount);
        subtitleRow.getCell(1).font = { size: 10, italic: true };
        subtitleRow.getCell(1).alignment = { vertical: 'middle' };
        subtitleRow.getCell(1).fill = headerFill;
        headerOffset += 1;
      }

      // 2 blank spacer rows between title block and column headers
      worksheet.addRow([]);
      worksheet.addRow([]);
      headerOffset += 2;
    }

    // Define columns without `header` — setting `header` here would overwrite
    // row 1 (the title row). We add the header row manually below instead.
    worksheet.columns = headers.map((h) => ({
      key: h,
      width: MIN_COL_WIDTH,
    }));

    // Manually add the column header row so it lands after any title rows
    const colHeaderFill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD9E1F2' }, // light blue-grey
    };
    const dataHeaderExcelRow = worksheet.addRow(headers);
    dataHeaderExcelRow.font = { bold: true };
    dataHeaderExcelRow.eachCell({ includeEmpty: true }, (cell) => {
      cell.fill = colHeaderFill;
    });

    const dataHeaderRow = dataHeaderExcelRow.number;


    // Add data rows
    rows.forEach((row) => {
      const values = headers.map((h) => {
        const v = row[h];
        return v !== undefined && v !== null ? v : '';
      });

      const excelRow = worksheet.addRow(values);

      // Format numeric cells as currency; enable wrap text for long strings
      excelRow.eachCell({ includeEmpty: true }, (cell) => {
        if (typeof cell.value === 'number') {
          cell.numFmt = '$#,##0.00';
        } else {
          const strVal = cell.value !== null && cell.value !== undefined
            ? String(cell.value)
            : '';
          if (strVal.length > WRAP_TEXT_THRESHOLD) {
            cell.alignment = { wrapText: true, vertical: 'top' };
          }
        }
      });
    });

    // Auto-size columns based on the longest value in each column
    // (including the header text), capped at MAX_COL_WIDTH.
    headers.forEach((_header, colIndex) => {
      const column = worksheet.getColumn(colIndex + 1);
      let maxLen = 0;

      column.eachCell({ includeEmpty: false }, (cell) => {
        // Skip merged title/subtitle rows — their text would inflate column 1
        if (cell.row <= headerOffset) return;
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
