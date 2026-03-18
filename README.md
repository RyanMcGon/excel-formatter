# excel-formatter

A stateless HTTP microservice that accepts JSON from an **n8n HTTP Request node** and returns a formatted, downloadable `.xlsx` Excel file.

---

## Features

- `POST /generate-excel` — converts a JSON rows array to a formatted `.xlsx` workbook
- `GET /health` — health check for Render / Railway / uptime monitors
- Bold header row, frozen top row, autofilter, auto-sized columns, wrap text on long cells
- Correct `Content-Type` and `Content-Disposition` headers for binary download
- Structured JSON logging via [pino](https://getpino.io/)

---

## Install

```bash
npm install
```

---

## Run locally

```bash
node index.js
```

The service starts on port `3000` by default. To use a different port:

```bash
PORT=8080 node index.js
```

Pretty-printed logs are enabled automatically in non-production environments.

---

## Test

```bash
npm test
```

---

## Request format

`POST /generate-excel`
`Content-Type: application/json`

```json
{
  "rows": [
    {
      "Record_ID": "a01fj00000azMN7AAM",
      "Field_Name": "Amount_Paid",
      "Old_Value": "3500",
      "New_Value": "4500"
    },
    {
      "Record_ID": "a01fj00000azMN8BBN",
      "Field_Name": "Status",
      "Old_Value": "Pending",
      "New_Value": "Approved"
    },
    {
      "Record_ID": "a01fj00000azMN9CCO",
      "Field_Name": "Notes",
      "Old_Value": "Short note",
      "New_Value": "This is a much longer note that exceeds the fifty-character wrap text threshold and should trigger wrapText in the Excel cell output."
    }
  ],
  "fileName": "salesforce_report.xlsx",
  "sheetName": "Report"
}
```

| Field | Type | Required | Default |
|---|---|---|---|
| `rows` | array of objects | Yes | — |
| `fileName` | string | No | `salesforce_report.xlsx` |
| `sheetName` | string | No | `Report` |

**Column ordering:** `Record_ID`, `Field_Name`, `Old_Value`, `New_Value` appear first (in that order). Any additional keys found in the row objects are appended after those four.

---

## Example curl request

Save the response directly as an `.xlsx` file:

```bash
curl -X POST https://your-service.onrender.com/generate-excel \
  -H "Content-Type: application/json" \
  -d '{
    "rows": [
      {
        "Record_ID": "a01fj00000azMN7AAM",
        "Field_Name": "Amount_Paid",
        "Old_Value": "3500",
        "New_Value": "4500"
      },
      {
        "Record_ID": "a01fj00000azMN8BBN",
        "Field_Name": "Status",
        "Old_Value": "Pending",
        "New_Value": "Approved"
      },
      {
        "Record_ID": "a01fj00000azMN9CCO",
        "Field_Name": "Notes",
        "Old_Value": "Short note",
        "New_Value": "This is a much longer note that exceeds the fifty-character wrap text threshold and should trigger wrapText in the Excel cell output."
      }
    ],
    "fileName": "salesforce_report.xlsx",
    "sheetName": "Report"
  }' \
  --output salesforce_report.xlsx
```

Local test (same command, different URL):

```bash
curl -X POST http://localhost:3000/generate-excel \
  -H "Content-Type: application/json" \
  -d '{ "rows": [{ "Record_ID": "001", "Field_Name": "Amount", "Old_Value": "100", "New_Value": "200" }] }' \
  --output test.xlsx && open test.xlsx
```

---

## Error responses

| Scenario | Status | Response |
|---|---|---|
| `rows` missing or not an array | 400 | `{ "error": "rows must be a non-empty array" }` |
| `rows` is empty | 400 | `{ "error": "rows must not be empty" }` |
| A row is not a plain object | 400 | `{ "error": "each row must be a plain object" }` |
| Payload exceeds 10 MB | 413 | Express default message |
| Internal Excel generation error | 500 | `{ "error": "Failed to generate Excel file" }` |

---

## Deploy to Render

This repo includes a `render.yaml` for one-click deployment.

### Steps

1. Push this project to a GitHub repository.
2. Go to [render.com](https://render.com) → **New** → **Web Service**.
3. Connect your GitHub repo.
4. Render detects `render.yaml` automatically and pre-fills all settings.
5. Click **Create Web Service**.

Your service URL will be something like `https://excel-formatter.onrender.com`.

### Manual settings (if not using render.yaml)

| Setting | Value |
|---|---|
| Environment | Node |
| Build Command | `npm install --omit=dev` |
| Start Command | `node index.js` |
| Health Check Path | `/health` |
| Environment Variable `NODE_ENV` | `production` |

---

## Keeping the service warm (free tier)

Render's free tier puts services to sleep after **15 minutes of inactivity**. The first n8n request after a sleep period can take 5–30 seconds to respond while the service cold-starts.

**Fix: use a free external pinger.**

1. Go to [cron-job.org](https://cron-job.org) (free account).
2. Create a new cron job:
   - **URL:** `https://your-service.onrender.com/health`
   - **Schedule:** Every 10 minutes
3. Save. The service stays warm and n8n requests will respond immediately.

---

## n8n HTTP Request node configuration

Configure the **HTTP Request** node in your n8n workflow as follows:

| Setting | Value |
|---|---|
| **Method** | POST |
| **URL** | `https://your-service.onrender.com/generate-excel` |
| **Authentication** | None |
| **Send Body** | ✅ Enabled |
| **Body Content Type** | JSON |
| **Response Format** | **File** |

**Body (JSON):**
```json
{
  "rows": {{ $json.rows }},
  "fileName": "salesforce_report.xlsx",
  "sheetName": "Report"
}
```

> Replace `{{ $json.rows }}` with the expression that references your row data from a previous node. You can also hardcode a static array for testing.

**Why "File" response format?**
Setting Response Format to **File** tells n8n to treat the response body as binary data instead of trying to parse it as JSON. The `.xlsx` file will be available in subsequent nodes as a binary item (e.g., to attach to an email or upload to Google Drive).

**Tip:** After the HTTP Request node, connect an **Email** or **Google Drive** node and reference the binary output with `{{ $binary.data }}`.

---

## Optional: API key authentication

By default the endpoint is unauthenticated. It is secured only by keeping the URL private. This is appropriate for internal n8n workflows.

If you need to expose the service more broadly, add a simple Bearer token check:

1. Set the `API_KEY` environment variable in Render's dashboard.
2. Add this middleware to `index.js` **before** the route definitions:

```javascript
app.use((req, res, next) => {
  const key = process.env.API_KEY;
  if (!key) return next(); // auth disabled when env var is unset
  const auth = req.headers['authorization'] || '';
  if (auth !== `Bearer ${key}`) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  next();
});
```

3. In your n8n HTTP Request node, add a **Header**:
   - Name: `Authorization`
   - Value: `Bearer your-secret-key-here`

---

## Project structure

```
excel-formatter/
├── index.js          # Express server — single entry point
├── package.json
├── render.yaml       # Render deployment config
├── .env.example      # Environment variable reference
├── .gitignore
├── README.md
└── test/
    └── excel.test.js # Jest + Supertest integration tests
```
