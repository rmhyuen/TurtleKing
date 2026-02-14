# Deployment Information

## Quick Reference

### Apps Script Deployment URL
```
https://script.google.com/macros/s/AKfycbwknF5fBpZHwy-U3nIhlOA8nWyKKLRx48VfT87XaAithJ3BcpcVx3nIWcY4fXw21dxh/exec
```

### Google Drive Configuration

**Folder ID:** `1EWsCGFT1iqMwUad48t8V_n2Irko-sxGZ`

**Folder Structure:**
```
WEBSITE_APP/
├── CustomerData/
│   └── Beals/
│       ├── New/         ← Input: Customer PDFs go here
│       ├── Processed/
│       │   ├── CustomerDataOnly/        ← Output: Step 1 processed customer Excel
│       │   └── CustomerAndVendorData/   ← Output: Step 2 merged customer+vendor Excel
│       └── Old/         ← Archive: Processed PDFs moved here
└── VendorData/          ← Input: Vendor data file (CSV/Excel) - latest auto-fetched
```

---

## Automated Workflow

### Step 1: Import & Process Customer Data
1. Select folder with customer PO PDFs
2. App processes PDFs client-side
3. **Automatically uploads** to `CustomerData/Beals/Processed/CustomerDataOnly/`
4. Filename: `CustomerData_YYYYMMDD_HHMMSS.xlsx`
5. Optional: Save copy locally

### Step 2: Import & Process Vendor Data
1. Click "Import & Process Vendor Data"
2. **Automatically fetches** latest customer data from `CustomerDataOnly/`
3. **Automatically fetches** vendor file from `VendorData/`
4. Merges data client-side in browser
5. **Automatically uploads** to `CustomerData/Beals/Processed/CustomerAndVendorData/`
6. Filename: `CustomerAndVendorData_YYYYMMDD_HHMMSS.xlsx`
7. Optional: Save copy locally

**No manual file selection needed for Step 2!**

---

## How to Use (Updated Flow)

### Step 1: Process Customer Data
1. Open `index.html` in browser
2. Click "Import & Process Customer Data"
3. Select folder with customer PO PDFs
4. Files automatically upload to Google Drive with timestamp
5. Optionally save a local copy

### Step 2: Process Vendor Data
1. Ensure vendor file is in `VendorData/` folder on Google Drive
2. Click "Import & Process Vendor Data"
3. App automatically fetches both customer and vendor data
4. Merges and uploads result automatically
5. Optionally save a local copy

**No manual file selection needed for Step 2 - it's fully automated!**

---

## Maintenance

### To Update the Apps Script:
1. Go to https://script.google.com
2. Open "PO Processor API" project
3. Make changes
4. Save (Ctrl+S)
5. Deploy → Manage deployments → Click pencil icon → Version: New version → Deploy
6. URL stays the same (no need to update in app)

### To Change Folder Structure:
1. Edit the `FOLDER_STRUCTURE` object in the Apps Script
2. Save and redeploy (see above)

### If Getting Errors:
- **403 Forbidden:** Re-authorize the script (Run button in Apps Script editor)
- **Folder not found:** Check folder names match exactly (case-sensitive)
- **No files found:** Make sure files are in the correct `New/` folders

---

## Project Links

- **Apps Script Editor:** https://script.google.com
- **Google Drive Folder:** https://drive.google.com/drive/folders/1EWsCGFT1iqMwUad48t8V_n2Irko-sxGZ
- **Full Setup Guide:** [GOOGLE_APPS_SCRIPT_SETUP.md](./GOOGLE_APPS_SCRIPT_SETUP.md)

---

**Last Updated:** February 6, 2026  
**Status:** Active ✅
