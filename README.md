# Quotation Maker Pro

A client-ready quotation maker with Google Sheets sync, invoice fetch, and WhatsApp sharing.

## Features

- **Create quotations** with company details, client info, and itemized pricing
- **Simple/Detailed modes** — toggle to hide or show rate & item-wise amounts
- **Google Sheets sync** — save quotations and item catalog to Google Sheets
- **Invoice fetch** — search by invoice number, client name, or mobile number
- **WhatsApp sharing** — share quotations directly via WhatsApp
- **JPEG export** — download client-ready quotation images
- **Print/PDF** — print or save as PDF
- **Responsive design** — works on desktop and mobile

## Live Demo

Visit the [GitHub Pages site](https://<your-username>.github.io/quotation_maker/) to try it out.

## Files

| File | Description |
|------|-------------|
| `index.html` | Main quotation maker application |
| `google_sheet_apps_script_v2.gs` | Google Apps Script backend for Sheets integration |
| `QUOTATION_MAKER_COMPLETE_GUIDE.txt` | Complete setup and usage guide |

## Deployment

### GitHub Pages (Automatic)

This repository is configured with GitHub Actions to automatically deploy to GitHub Pages on every push to `main`.

### Manual Setup

1. **Fork or clone** this repository
2. Enable GitHub Pages in **Settings → Pages → Source: GitHub Actions**
3. The site will be deployed automatically

## Google Sheets Integration

1. Create a Google Sheet
2. Go to **Extensions → Apps Script**
3. Paste the contents of `google_sheet_apps_script_v2.gs`
4. Deploy as Web App (Execute as: Me, Access: Anyone)
5. Copy the `/exec` URL
6. Open the app → **Settings** → paste the URL → **Test Connection**

## Local Development

Simply open `index.html` in a browser. No build tools required.

## License

MIT
