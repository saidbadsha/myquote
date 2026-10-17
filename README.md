# Invoice & Quotation Maker Pro

A professional invoicing and quotation system with Google Sheets sync, auto-numbering, PDF/JPEG export, and WhatsApp sharing.

## Features

- **Invoice & Quotation** — Create both with separate numbering (INV-/QT-)
- **Auto-numbering** — Automatic next document number from Google Sheets
- **Google Sheets sync** — Invoices, quotations, company details, and item catalog stored in Sheets
- **Separate records** — Invoices and quotations are stored independently, never overwrite each other
- **Item catalog** — Previously saved item rates auto-fill when reused
- **Company details** — Saved to Sheet and auto-loaded on future sessions
- **JPEG export** — Download document images matching the preview
- **PDF export** — Direct PDF save via browser print dialog
- **WhatsApp sharing** — Share documents directly via WhatsApp
- **Simple/Detailed mode** — Toggle to hide or show rate & amounts
- **Responsive design** — Works on desktop and mobile
- **Professional UI** — Color-coded buttons for different actions

## Files

| File | Description |
|------|-------------|
| `index.html` | Main application (open in browser) |
| `admin.html` | Admin login + dashboard (invoices & quotations) |
| `google_sheet_apps_script_v2.gs` | Google Apps Script backend for Sheets (v4: numbering, admin auth) |
| `tests/` | `node tests/backend.test.js` runs the backend against an in-memory Sheets mock |
| `QUOTATION_MAKER_COMPLETE_GUIDE.txt` | Complete setup and usage guide |

## Quick Start

1. Open `index.html` in a browser
2. Go to **Settings** and paste your Google Apps Script Web App URL
3. Click **Test Connection** then **Save Company to Sheet**
4. Select **Invoice** or **Quotation** at the top
5. Start creating documents

## Google Sheets Integration

1. Create a Google Sheet
2. **Extensions → Apps Script**
3. Paste contents of `google_sheet_apps_script_v2.gs`
4. Deploy as Web App (Execute as: Me, Access: Anyone)
5. Copy the `/exec` URL and paste in **Settings**

### Admin setup (one time)

1. In the Apps Script editor run **`setupAdmin`** and read the password from the Execution log (shown once, never stored in plain text; re-run to rotate it and sign out all sessions).
2. Sign in at `admin.html` with `admin@email.com`.
3. Optional: set Script Property `ADMIN_COMPANIES` to `Company A|Company B` to limit the admin to those companies (default `*` = all).

### Company-wise numbering

**Settings → Document Numbering** sets invoice/quotation formats per company, e.g. `ACM/{YY}{MM}/{SEQ:4}`. Default stays `INV-{YYYY}-{SEQ:3}` / `QT-{YYYY}-{SEQ:3}`. Existing numbers are never changed; the server assigns the final number on save under a lock.

The script auto-creates: `CompanyDetails`, `Invoices`, `Quotations`, `ItemCatalog`, `DocCounters`

## Deployment

### GitHub Pages

Push to `main` branch — auto-deploys via GitHub Actions.

## License

MIT
