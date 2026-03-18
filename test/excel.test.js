'use strict';

const request = require('supertest');
const ExcelJS = require('exceljs');
const app = require('../index');

// ─── Helpers ──────────────────────────────────────────────────────────────────

async function parseXlsx(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return workbook;
}

function getHeaders(worksheet) {
  const headers = [];
  worksheet.getRow(1).eachCell({ includeEmpty: false }, (cell) => {
    headers.push(cell.value);
  });
  return headers;
}

// ─── Sample data ──────────────────────────────────────────────────────────────

const SAMPLE_ROWS = [
  {
    Record_ID: 'a01fj00000azMN7AAM',
    Field_Name: 'Amount_Paid',
    Old_Value: '3500',
    New_Value: '4500',
  },
  {
    Record_ID: 'a01fj00000azMN8BBN',
    Field_Name: 'Status',
    Old_Value: 'Pending',
    New_Value: 'Approved',
  },
  {
    Record_ID: 'a01fj00000azMN9CCO',
    Field_Name: 'Notes',
    Old_Value: 'Short note',
    New_Value:
      'This is a much longer note that exceeds the fifty-character wrap text threshold and should trigger wrapText in the Excel cell output.',
  },
];

// ─── Tests ────────────────────────────────────────────────────────────────────

describe('GET /health', () => {
  it('returns 200 with ok status', async () => {
    const res = await request(app).get('/health');
    expect(res.status).toBe(200);
    expect(res.body.status).toBe('ok');
    expect(typeof res.body.ts).toBe('string');
  });
});

describe('POST /generate-excel — happy path', () => {
  it('returns 200 with xlsx content-type for a valid payload', async () => {
    const res = await request(app)
      .post('/generate-excel')
      .send({ rows: SAMPLE_ROWS })
      .buffer(true)
      .parse((res, cb) => {
        const chunks = [];
        res.on('data', (c) => chunks.push(c));
        res.on('end', () => cb(null, Buffer.concat(chunks)));
      });

    expect(res.status).toBe(200);
    expect(res.headers['content-type']).toMatch(
      /application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet/
    );
  });

  it('sets Content-Disposition with provided fileName', async () => {
    const res = await request(app)
      .post('/generate-excel')
      .send({ rows: SAMPLE_ROWS, fileName: 'my_report.xlsx' })
      .buffer(true)
      .parse((res, cb) => {
        const chunks = [];
        res.on('data', (c) => chunks.push(c));
        res.on('end', () => cb(null, Buffer.concat(chunks)));
      });

    expect(res.headers['content-disposition']).toContain('my_report.xlsx');
  });

  it('defaults fileName to salesforce_report.xlsx when omitted', async () => {
    const res = await request(app)
      .post('/generate-excel')
      .send({ rows: SAMPLE_ROWS })
      .buffer(true)
      .parse((res, cb) => {
        const chunks = [];
        res.on('data', (c) => chunks.push(c));
        res.on('end', () => cb(null, Buffer.concat(chunks)));
      });

    expect(res.headers['content-disposition']).toContain('salesforce_report.xlsx');
  });

  it('returns a non-empty buffer', async () => {
    const res = await request(app)
      .post('/generate-excel')
      .send({ rows: SAMPLE_ROWS })
      .buffer(true)
      .parse((res, cb) => {
        const chunks = [];
        res.on('data', (c) => chunks.push(c));
        res.on('end', () => cb(null, Buffer.concat(chunks)));
      });

    expect(res.body.length).toBeGreaterThan(0);
  });
});

describe('POST /generate-excel — column ordering', () => {
  let worksheet;

  beforeAll(async () => {
    const res = await request(app)
      .post('/generate-excel')
      .send({ rows: SAMPLE_ROWS })
      .buffer(true)
      .parse((res, cb) => {
        const chunks = [];
        res.on('data', (c) => chunks.push(c));
        res.on('end', () => cb(null, Buffer.concat(chunks)));
      });

    const workbook = await parseXlsx(res.body);
    worksheet = workbook.getWorksheet(1);
  });

  it('places Record_ID as the first column', () => {
    const headers = getHeaders(worksheet);
    expect(headers[0]).toBe('Record_ID');
  });

  it('places Field_Name as the second column', () => {
    const headers = getHeaders(worksheet);
    expect(headers[1]).toBe('Field_Name');
  });

  it('places Old_Value as the third column', () => {
    const headers = getHeaders(worksheet);
    expect(headers[2]).toBe('Old_Value');
  });

  it('places New_Value as the fourth column', () => {
    const headers = getHeaders(worksheet);
    expect(headers[3]).toBe('New_Value');
  });

  it('appends extra keys after the four priority headers', async () => {
    const rows = [{ Record_ID: '001', Field_Name: 'F', Old_Value: 'a', New_Value: 'b', ExtraCol: 'x' }];
    const res = await request(app)
      .post('/generate-excel')
      .send({ rows })
      .buffer(true)
      .parse((res, cb) => {
        const chunks = [];
        res.on('data', (c) => chunks.push(c));
        res.on('end', () => cb(null, Buffer.concat(chunks)));
      });

    const workbook = await parseXlsx(res.body);
    const ws = workbook.getWorksheet(1);
    const headers = getHeaders(ws);
    expect(headers[4]).toBe('ExtraCol');
  });

  it('includes columns from all rows when rows have heterogeneous keys', async () => {
    const rows = [
      { Record_ID: '001', Field_Name: 'F', Old_Value: 'a', New_Value: 'b', ColA: '1' },
      { Record_ID: '002', Field_Name: 'G', Old_Value: 'c', New_Value: 'd', ColB: '2' },
    ];
    const res = await request(app)
      .post('/generate-excel')
      .send({ rows })
      .buffer(true)
      .parse((res, cb) => {
        const chunks = [];
        res.on('data', (c) => chunks.push(c));
        res.on('end', () => cb(null, Buffer.concat(chunks)));
      });

    const workbook = await parseXlsx(res.body);
    const ws = workbook.getWorksheet(1);
    const headers = getHeaders(ws);
    expect(headers).toContain('ColA');
    expect(headers).toContain('ColB');
  });
});

describe('POST /generate-excel — Excel formatting', () => {
  let worksheet;

  beforeAll(async () => {
    const res = await request(app)
      .post('/generate-excel')
      .send({ rows: SAMPLE_ROWS })
      .buffer(true)
      .parse((res, cb) => {
        const chunks = [];
        res.on('data', (c) => chunks.push(c));
        res.on('end', () => cb(null, Buffer.concat(chunks)));
      });

    const workbook = await parseXlsx(res.body);
    worksheet = workbook.getWorksheet(1);
  });

  it('makes the header row bold', () => {
    const headerCell = worksheet.getRow(1).getCell(1);
    expect(headerCell.font?.bold).toBe(true);
  });

  it('freezes the top row', () => {
    const views = worksheet.views;
    expect(views.length).toBeGreaterThan(0);
    expect(views[0].state).toBe('frozen');
    expect(views[0].ySplit).toBe(1);
  });

  it('applies autofilter to the header row', () => {
    expect(worksheet.autoFilter).toBeTruthy();
  });

  it('sets column widths based on content (wider than minimum)', () => {
    // Record_ID values are long — column should be wider than MIN_COL_WIDTH
    const col = worksheet.getColumn(1);
    expect(col.width).toBeGreaterThan(10);
  });

  it('caps column width at MAX_COL_WIDTH (50)', () => {
    // All columns should be <= 50
    worksheet.columns.forEach((col) => {
      if (col.width !== undefined) {
        expect(col.width).toBeLessThanOrEqual(50);
      }
    });
  });

  it('enables wrap text for cells with content longer than 50 characters', () => {
    // Row 4 = data row 3 (the long-text row), which is worksheet row 4
    const longCell = worksheet.getRow(4).getCell(4); // New_Value column
    expect(longCell.alignment?.wrapText).toBe(true);
  });

  it('uses the provided sheetName for the worksheet', async () => {
    const res = await request(app)
      .post('/generate-excel')
      .send({ rows: SAMPLE_ROWS, sheetName: 'MySheet' })
      .buffer(true)
      .parse((res, cb) => {
        const chunks = [];
        res.on('data', (c) => chunks.push(c));
        res.on('end', () => cb(null, Buffer.concat(chunks)));
      });

    const workbook = await parseXlsx(res.body);
    expect(workbook.getWorksheet('MySheet')).toBeTruthy();
  });

  it('defaults sheetName to Report when omitted', async () => {
    const res = await request(app)
      .post('/generate-excel')
      .send({ rows: SAMPLE_ROWS })
      .buffer(true)
      .parse((res, cb) => {
        const chunks = [];
        res.on('data', (c) => chunks.push(c));
        res.on('end', () => cb(null, Buffer.concat(chunks)));
      });

    const workbook = await parseXlsx(res.body);
    expect(workbook.getWorksheet('Report')).toBeTruthy();
  });
});

describe('POST /generate-excel — validation (400s)', () => {
  const cases = [
    ['rows is missing', {}],
    ['rows is null', { rows: null }],
    ['rows is a string', { rows: 'bad' }],
    ['rows is an object (not array)', { rows: { a: 1 } }],
    ['rows is a number', { rows: 42 }],
    ['rows is an empty array', { rows: [] }],
    ['a row is null', { rows: [null] }],
    ['a row is a primitive string', { rows: ['bad'] }],
  ];

  test.each(cases)('returns 400 when %s', async (_label, body) => {
    const res = await request(app).post('/generate-excel').send(body);
    expect(res.status).toBe(400);
    expect(res.body.error).toBeTruthy();
  });
});

describe('POST /generate-excel — sheetName sanitization', () => {
  it('sanitizes sheetName containing invalid Excel characters', async () => {
    const res = await request(app)
      .post('/generate-excel')
      .send({ rows: SAMPLE_ROWS, sheetName: 'My[Sheet]' })
      .buffer(true)
      .parse((res, cb) => {
        const chunks = [];
        res.on('data', (c) => chunks.push(c));
        res.on('end', () => cb(null, Buffer.concat(chunks)));
      });

    expect(res.status).toBe(200);
    const workbook = await parseXlsx(res.body);
    // Invalid chars stripped — worksheet name should exist (sanitized form)
    expect(workbook.worksheets.length).toBe(1);
  });

  it('truncates sheetName longer than 31 characters', async () => {
    const longName = 'A'.repeat(40);
    const res = await request(app)
      .post('/generate-excel')
      .send({ rows: SAMPLE_ROWS, sheetName: longName })
      .buffer(true)
      .parse((res, cb) => {
        const chunks = [];
        res.on('data', (c) => chunks.push(c));
        res.on('end', () => cb(null, Buffer.concat(chunks)));
      });

    expect(res.status).toBe(200);
    const workbook = await parseXlsx(res.body);
    const ws = workbook.worksheets[0];
    expect(ws.name.length).toBeLessThanOrEqual(31);
  });
});
