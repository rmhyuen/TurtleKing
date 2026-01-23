# PO Data Processor

A 100% client-side web application for processing Purchase Order PDFs and merging vendor data. All processing happens in your browser - your data never leaves your computer.

## 🔒 Privacy First

- **No server required** - Everything runs in your browser
- **Your data stays local** - Files are never uploaded anywhere
- **Works offline** - Once loaded, no internet connection needed

## Recent Updates

- **Fixed Pack Qty extraction** for complex size descriptors (e.g., "NO SIZE", "ONE SIZE")
- Improved regex pattern matching for various PO formats
- Better handling of multi-line format extraction

## Features

### 1. Import & Process Customer Data
- Select a folder containing PO PDF files
- Extracts SKU data, pricing, quantities from PDFs
- Merges data from multiple PDFs by SKU
- Generates a formatted Excel file with:
  - SKU and MFG Style columns
  - Cost/Unit and Retail pricing
  - Total Amount and Total Units with formulas
  - Pack Qty and per-PO quantity columns

### 2. Import & Process Vendor Data
- Takes the output from step 1 (Customer Excel)
- Merges with vendor data (CSV or Excel)
- Adds vendor-specific columns:
  - Vend # (Vendor Number)
  - Base Cost
  - Box/Case
  - Unit/Case
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
