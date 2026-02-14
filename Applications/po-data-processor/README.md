# PO Data Processor

A 100% client-side web application for processing Purchase Order PDFs and merging vendor data. All processing happens in your browser - your data never leaves your computer.

## Features

### 1. Import & Process Customer Data
- Select a folder containing PO PDF files
- Extracts SKU data, pricing, quantities from PDFs
- Merges data from multiple PDFs by SKU
- Automatically uploads processed Excel to Google Drive:
  - Destination: `CustomerData/Beals/Processed/CustomerDataOnly/`
  - Filename: `CustomerData_YYYYMMDD_HHMMSS.xlsx` (with automatic timestamp)
- Optionally saves a copy locally as well
- Generates a formatted Excel file with:
  - SKU and MFG Style columns
  - Cost/Unit and Retail pricing
  - Total Amount and Total Units with formulas
  - Pack Qty and per-PO quantity columns

### 2. Import & Process Vendor Data
- **Fully automated workflow** - No manual file selection needed!
- Automatically fetches the latest processed customer data from Google Drive
- Automatically fetches the vendor data file from Google Drive
- Merges both files client-side in your browser
- Adds vendor-specific columns:
  - Vend # (Vendor Number)
  - Base Cost
  - Box/Case
  - Unit/Case
- Automatically uploads merged result to Google Drive:
  - Destination: `CustomerData/Beals/Processed/CustomerAndVendorData/`
  - Filename: `CustomerAndVendorData_YYYYMMDD_HHMMSS.xlsx` (with automatic timestamp)
- Optionally saves a copy locally as well
- Automatically adjusts formulas for new columns

## Usage

### Option 1: GitHub Pages (Recommended)
Visit the live site: [https://rmhyuen.github.io/TurtleKing/apps/po-data-processor](https://rmhyuen.github.io/TurtleKing/apps/po-data-processor)

### Option 2: Run Locally
Simply open `index.html` in a modern web browser (Chrome, Firefox, Edge, Safari).

No installation or server setup required!

## How It Works

1. **PDF.js** (Mozilla) - Extracts text from PDF files
2. **ExcelJS** - Creates formatted Excel files with formulas
3. **Custom Parsing** - Handles multiple PO PDF formats with 4 extraction strategies

## Browser Compatibility

Works in all modern browsers:
- Chrome 90+
- Firefox 88+
- Edge 90+
- Safari 14+

## 🔗 Google Drive Integration Setup

The app can automatically upload processed customer data to Google Drive. To enable this feature:

### Step 1: Deploy Google Apps Script
1. Go to [script.google.com](https://script.google.com)
2. Create a new project
3. Copy the entire contents of `GoogleAppsScript_Template.example.gs` into the editor
4. Update the `AUTHORIZED_API_KEY` constant with your own secure key (or keep the default)
5. Update the `FOLDER_ID` on line 27 with your folder ID from:
   - `https://drive.google.com/drive/folders/1EWsCGFT1iqMwUad48t8V_n2Irko-sxGZ`
   - Extract: `1EWsCGFT1iqMwUad48t8V_n2Irko-sxGZ`
6. Save the project
7. Click **Deploy** → **New Deployment**
8. Select type: **Web app**
9. Execute as: Your Google account
10. Who has access: **Anyone**
11. Click **Deploy** (the deployment URL is already configured in the app)

### Step 2: Configure API Key on First Visit
1. Open the PO Data Processor app
2. A setup modal appears automatically on first visit
3. Enter your API key (the `AUTHORIZED_API_KEY` from Step 1)
4. Click **"Save & Continue"**
5. Your API key is securely stored in your browser (never sent anywhere except to the Apps Script backend)

**That's it!** The app now has access to your Google Drive. You can reconfigure the API key anytime by clicking the settings button (⚙️) in the top right.

### Step 3: Set Up Folder Structure
Ensure your Google Drive has this folder structure:
```
CustomerData/
└── Beals/
    ├── New/              (Source folder for input PDFs)
    └── Processed/
        ├── CustomerDataOnly/        (Step 1 auto-upload destination)
        └── CustomerAndVendorData/   (Step 2 auto-upload destination)
VendorData/
    (Place vendor Excel/CSV file here - latest file will be auto-fetched)
```

The application now provides **full automation**:

**Step 1 - Customer Data:**
- Upload processed files to `CustomerDataOnly/`
- Name files as: `CustomerData_YYYYMMDD_HHMMSS.xlsx`
- Display upload status to the user

**Step 2 - Vendor Merge:**
- Auto-fetch latest customer data from `CustomerDataOnly/`
- Auto-fetch vendor data from `VendorData/`
- Merge files client-side
- Auto-upload result to `CustomerAndVendorData/`
- Name files as: `CustomerAndVendorData_YYYYMMDD_HHMMSS.xlsx`
- No manual file selection needed!

## Deploying to GitHub Pages

The app is part of the TurtleKing repository and automatically deploys to GitHub Pages. Simply push changes to the main branch:

```bash
git add .
git commit -m "Update: your changes"
git push origin main
```

Your site will be live at `https://rmhyuen.github.io/TurtleKing/apps/po-data-processor`

## File Structure

```
po-data-processor/
├── index.html          # Main HTML page
├── css/
│   └── style.css       # Styles
├── js/
│   ├── app.js          # Application logic
│   └── converter.js    # PDF parsing and Excel generation
└── README.md           # This file
```

## License

MIT License
