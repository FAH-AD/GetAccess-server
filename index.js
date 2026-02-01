/**
 * Express server: POST /api/subscribe writes email to Google Sheet.
 * Uses dotenv for local dev; in Render, set env vars in dashboard.
 */

import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import { google } from 'googleapis';

const app = express();
const PORT = process.env.PORT ?? 3001;

// -----------------------------------------------------------------------------
// Middleware
// -----------------------------------------------------------------------------

// Optional: lock CORS to a single frontend origin on Render
// Set ALLOWED_ORIGIN=https://your-frontend-domain.com in Render
const allowedOrigin = process.env.ALLOWED_ORIGIN;
app.use(
  cors(
    allowedOrigin
      ? { origin: allowedOrigin, methods: ['GET', 'POST', 'OPTIONS'], credentials: false }
      : undefined
  )
);

app.use(express.json());

// -----------------------------------------------------------------------------
// Google Sheets client (Service Account) - ENV JSON (Render-friendly)
// -----------------------------------------------------------------------------

/**
 * Requires:
 * - GOOGLE_SHEET_ID
 * - GOOGLE_SERVICE_ACCOUNT_JSON (the full JSON contents)
 */
function loadServiceAccountFromEnv() {
  const sheetId = process.env.GOOGLE_SHEET_ID;
  if (!sheetId?.trim()) throw new Error('GOOGLE_SHEET_ID is missing');

  const rawJson = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  if (!rawJson?.trim()) throw new Error('GOOGLE_SERVICE_ACCOUNT_JSON is missing');

  let key;
  try {
    key = JSON.parse(rawJson);
  } catch {
    throw new Error('GOOGLE_SERVICE_ACCOUNT_JSON is invalid JSON');
  }

  // Common gotcha: private_key may contain escaped newlines in env vars
  if (typeof key.private_key === 'string') {
    key.private_key = key.private_key.replace(/\\n/g, '\n');
  }

  return { key, sheetId };
}

// Cache client so we don't parse JSON/create auth for every request
let cached = null;

function getSheetsClient() {
  if (cached) return cached;

  const { key, sheetId } = loadServiceAccountFromEnv();
  const auth = new google.auth.GoogleAuth({
    credentials: key,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  const sheets = google.sheets({ version: 'v4', auth });
  cached = { sheets, sheetId };
  return cached;
}

/** Returns the first sheet's internal ID (for batchUpdate). */
async function getFirstSheetId(sheets, spreadsheetId) {
  const { data } = await sheets.spreadsheets.get({ spreadsheetId });
  const first = data.sheets?.[0]?.properties;
  if (!first) throw new Error('No sheets found');
  return first.sheetId;
}

/**
 * Find row index (0-based in the values array) where column C (Email) matches.
 * Returns -1 if not found.
 *
 * Note: values.get includes header row at index 0
 */
async function findEmailRow(sheets, spreadsheetId, email) {
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: 'Sheet1!A:C',
  });

  const rows = data.values ?? [];
  const normalized = email.trim().toLowerCase();

  for (let i = 1; i < rows.length; i++) {
    const cell = (rows[i][2] ?? '').toString().trim().toLowerCase();
    if (cell === normalized) return i; // i is the row index in the sheet values (header is 0)
  }
  return -1;
}

/** Delete a single row (0-based index in the SHEET grid, row 0 = header). */
async function deleteSheetRow(sheets, spreadsheetId, gridSheetId, rowIndex0Based) {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          deleteDimension: {
            range: {
              sheetId: gridSheetId,
              dimension: 'ROWS',
              startIndex: rowIndex0Based,
              endIndex: rowIndex0Based + 1,
            },
          },
        },
      ],
    },
  });
}

/**
 * Appends a single row to Sheet1.
 * Columns: A = Status, B = Name, C = Email.
 */
async function appendRowToSheet(name, email) {
  const { sheets, sheetId } = getSheetsClient();

  await sheets.spreadsheets.values.append({
    spreadsheetId: sheetId,
    range: 'Sheet1!A:C',
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: {
      values: [['active', name, email]],
    },
  });
}

/**
 * Update-only flow: add only if email already exists.
 * If exists: delete that row, then append new row [active, name, email].
 * If not: throw (caller returns 400).
 */
async function updateExistingEmailOnly(name, email) {
  const { sheets, sheetId } = getSheetsClient();

  const rowIndex = await findEmailRow(sheets, sheetId, email);
  if (rowIndex === -1) throw new Error('EMAIL_NOT_FOUND');

  const gridSheetId = await getFirstSheetId(sheets, sheetId);
  await deleteSheetRow(sheets, sheetId, gridSheetId, rowIndex);
  await appendRowToSheet(name, email);
}

// -----------------------------------------------------------------------------
// Routes
// -----------------------------------------------------------------------------

const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

app.post('/api/subscribe', async (req, res) => {
  try {
    const name = typeof req.body?.name === 'string' ? req.body.name.trim() : '';
    const email = typeof req.body?.email === 'string' ? req.body.email.trim() : '';

    if (!name) return res.status(400).json({ ok: false, error: 'Name is required' });
    if (!email) return res.status(400).json({ ok: false, error: 'Email is required' });
    if (!EMAIL_REGEX.test(email))
      return res.status(400).json({ ok: false, error: 'Please enter a valid email address' });

    await updateExistingEmailOnly(name, email);

    return res.status(200).json({
      ok: true,
      message: 'Access is granted. You will shortly receive an email with your access link.',
    });
  } catch (err) {
    console.error('Subscribe error:', err?.message || err);

    const failureMessage =
      'We faced a problem while granting access. Kindly ensure you are already registered to our Association.';

    if (err?.message === 'EMAIL_NOT_FOUND') {
      return res.status(400).json({ ok: false, error: failureMessage });
    }

    return res.status(500).json({ ok: false, error: failureMessage });
  }
});

// -----------------------------------------------------------------------------
// Start
// -----------------------------------------------------------------------------

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});