/**
 * Express server: POST /api/subscribe writes email to Google Sheet.
 * Uses dotenv for env vars, CORS enabled, Google Sheets API v4 via Service Account.
 */

import 'dotenv/config';
import { readFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import express from 'express';
import cors from 'cors';
import { google } from 'googleapis';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();
const PORT = process.env.PORT ?? 3001;

// -----------------------------------------------------------------------------
// Middleware
// -----------------------------------------------------------------------------

app.use(cors());
app.use(express.json());

// -----------------------------------------------------------------------------
// Google Sheets client (Service Account)
// -----------------------------------------------------------------------------

/**
 * Load Service Account key from env.
 * Use either:
 *   - GOOGLE_SERVICE_ACCOUNT_FILE: path to key JSON (relative to server/ or absolute)
 *   - GOOGLE_SERVICE_ACCOUNT_JSON: full key JSON as string
 * Also requires GOOGLE_SHEET_ID.
 */
function loadServiceAccountKey() {
  const sheetId = process.env.GOOGLE_SHEET_ID;
  if (!sheetId?.trim()) {
    throw new Error('GOOGLE_SHEET_ID is missing');
  }

  const filePath = process.env.GOOGLE_SERVICE_ACCOUNT_FILE;
  if (filePath?.trim()) {
    const resolved = path.isAbsolute(filePath)
      ? filePath
      : path.join(__dirname, filePath);
    try {
      const raw = readFileSync(resolved, 'utf8');
      return { key: JSON.parse(raw), sheetId };
    } catch (e) {
      if (e.code === 'ENOENT') {
        throw new Error(`GOOGLE_SERVICE_ACCOUNT_FILE not found: ${filePath}`);
      }
      throw new Error('GOOGLE_SERVICE_ACCOUNT_FILE is invalid JSON');
    }
  }

  const raw = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  if (!raw?.trim()) {
    throw new Error(
      'Set GOOGLE_SERVICE_ACCOUNT_FILE (path to key JSON) or GOOGLE_SERVICE_ACCOUNT_JSON (JSON string)'
    );
  }
  try {
    return { key: JSON.parse(raw), sheetId };
  } catch (e) {
    throw new Error('GOOGLE_SERVICE_ACCOUNT_JSON is invalid JSON');
  }
}

function getSheetsClient() {
  const { key, sheetId } = loadServiceAccountKey();
  const auth = new google.auth.GoogleAuth({
    credentials: key,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return { auth, sheetId };
}

/** Returns the first sheet's internal ID (for batchUpdate). */
async function getFirstSheetId(sheets, spreadsheetId) {
  const { data } = await sheets.spreadsheets.get({ spreadsheetId });
  const first = data.sheets?.[0]?.properties;
  if (!first) throw new Error('No sheets found');
  return first.sheetId;
}

/**
 * Find 0-based row index (excluding header) where column C (Email) matches.
 * Returns -1 if not found.
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
    if (cell === normalized) return i;
  }
  return -1;
}

/** Delete a single row (0-based index; row 0 = header, row 1 = first data row). */
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
 * Appends a single row to the first sheet.
 * Columns: A = Status, B = Name (given name), C = Email.
 * Status is "active" by default for new rows.
 */
async function appendRowToSheet(name, email) {
  const { auth, sheetId } = getSheetsClient();
  const sheets = google.sheets({ version: 'v4', auth });

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
  const { auth, sheetId } = getSheetsClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const rowIndex = await findEmailRow(sheets, sheetId, email);
  if (rowIndex === -1) {
    throw new Error('EMAIL_NOT_FOUND');
  }

  const gridSheetId = await getFirstSheetId(sheets, sheetId);
  await deleteSheetRow(sheets, sheetId, gridSheetId, rowIndex);
  await appendRowToSheet(name, email);
}

// -----------------------------------------------------------------------------
// Routes
// -----------------------------------------------------------------------------

/** Simple email regex for basic validation */
const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

app.post('/api/subscribe', async (req, res) => {
  try {
    const name = typeof req.body?.name === 'string' ? req.body.name.trim() : '';
    const email = typeof req.body?.email === 'string' ? req.body.email.trim() : '';

    if (!name) {
      return res.status(400).json({ ok: false, error: 'Name is required' });
    }
    if (!email) {
      return res.status(400).json({ ok: false, error: 'Email is required' });
    }
    if (!EMAIL_REGEX.test(email)) {
      return res.status(400).json({ ok: false, error: 'Please enter a valid email address' });
    }

    await updateExistingEmailOnly(name, email);
    return res
      .status(200)
      .json({ ok: true, message: 'Access is granted. You will shortly receive an email with your access link.' });
  } catch (err) {
    console.error('Subscribe error:', err.message);
    const failureMessage =
      'We faced a problem while granting access. Kindly ensure you are already registered to our Association.';
    if (err.message === 'EMAIL_NOT_FOUND') {
      return res.status(400).json({ ok: false, error: failureMessage });
    }
    return res.status(500).json({ ok: false, error: failureMessage });
  }
});

// -----------------------------------------------------------------------------
// Start
// -----------------------------------------------------------------------------

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
