/**
 * PO Data Processor - Client-Side Application
 * 100% browser-based - no server required, works on GitHub Pages
 */

const customerBtn = document.getElementById('customer-btn');
const vendorBtn = document.getElementById('vendor-btn');
const fileInput = document.getElementById('file-input');
const customerExcelInput = document.getElementById('customer-excel-input');
const vendorDataInput = document.getElementById('vendor-data-input');
const statusEl = document.getElementById('status');
const resultPanel = document.getElementById('result-panel');
const resultLink = document.getElementById('result-link');

let processedBlob = null;
let processedFileName = 'merged-output.xlsx';

// State for vendor merge workflow
let vendorMergeState = {
  customerFile: null,
  vendorFile: null
};

/**
 * LocalStorage helpers for remembering last used paths
 */
const PathMemory = {
  keys: {
    lastPdfFolder: 'po-processor-last-pdf-folder',
    lastCustomerExcel: 'po-processor-last-customer-excel-folder',
    lastVendorData: 'po-processor-last-vendor-data-folder',
    lastOutputFolder: 'po-processor-last-output-folder'
  },
  
  save(key, path) {
    try {
      localStorage.setItem(key, path);
    } catch (e) {
      // localStorage might be disabled, silently fail
    }
  },
  
  get(key) {
    try {
      return localStorage.getItem(key);
    } catch (e) {
      return null;
    }
  }
};

/**
 * Show status message
 */
function showStatus(message, type = 'info') {
  statusEl.textContent = message;
  statusEl.classList.remove('info', 'error', 'success', 'loading');
  statusEl.classList.add('show', type);
}

function clearOutputButton() {
  resultPanel.classList.remove('show');
  resultLink.innerHTML = '';
}

/**
 * Handle customer data button click - open folder picker for INPUT
 */
customerBtn.addEventListener('click', () => {
  showStatus('Select folder with PDF files.', 'info');
  fileInput.click();
});

/**
 * Handle vendor data button click - start vendor merge workflow
 */
vendorBtn.addEventListener('click', () => {
  // Reset vendor merge state
  vendorMergeState = { customerFile: null, vendorFile: null };
  showStatus('Step 1/2: Select Customer Excel file.', 'info');
  customerExcelInput.click();
});

/**
 * Handle customer Excel file selection for vendor merge
 */
customerExcelInput.addEventListener('change', () => {
  if (!customerExcelInput.files || customerExcelInput.files.length === 0) {
    customerExcelInput.value = '';
    return;
  }
  
  vendorMergeState.customerFile = customerExcelInput.files[0];
  
  // Remember this folder
  PathMemory.save(PathMemory.keys.lastCustomerExcel, 'Customer Excel folder');
  
  customerExcelInput.value = '';
  
  // Show button for step 2 - browser requires direct user click for file picker
  showStatus(`Step 2/2: Now select the Vendor Data file`, 'info');
  resultLink.innerHTML = `
    <button id="select-vendor-btn" style="
      background: #2196F3;
      color: white;
      border: none;
      padding: 12px 24px;
      border-radius: 4px;
      cursor: pointer;
      font-weight: 600;
      font-size: 16px;
      width: 100%;
    ">Select Vendor Data File (CSV or Excel)</button>
  `;
  resultPanel.classList.add('show');
  
  document.getElementById('select-vendor-btn').addEventListener('click', () => {
    vendorDataInput.click();
  });
});

/**
 * Handle vendor data file selection for vendor merge (CLIENT-SIDE)
 */
vendorDataInput.addEventListener('change', async () => {
  if (!vendorDataInput.files || vendorDataInput.files.length === 0) {
    vendorDataInput.value = '';
    return;
  }
  
  vendorMergeState.vendorFile = vendorDataInput.files[0];
  vendorDataInput.value = '';
  
  // Remember folder path
  if (vendorMergeState.vendorFile) {
    PathMemory.save(PathMemory.keys.lastVendorData, 'Vendor data folder');
  }
  
  // Validate we have both files
  if (!vendorMergeState.customerFile || !vendorMergeState.vendorFile) {
    showStatus('Error: Missing files. Please try again.', 'error');
    return;
  }
  
  // Process the vendor merge CLIENT-SIDE
  showStatus(`Processing vendor merge... Please wait.`, 'loading');
  resultPanel.classList.remove('show');
  customerBtn.disabled = true;
  vendorBtn.disabled = true;
  
  try {
    // Read files as ArrayBuffer
    const customerBuffer = await vendorMergeState.customerFile.arrayBuffer();
    const vendorBuffer = await vendorMergeState.vendorFile.arrayBuffer();
    
    // Use client-side converter
    const mergedWorkbook = await window.POConverter.mergeVendorData(
      customerBuffer,
      vendorBuffer,
      vendorMergeState.vendorFile.name
    );
    
    // Convert workbook to blob
    const buffer = await mergedWorkbook.xlsx.writeBuffer();
    processedBlob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    
    // Generate output filename
    const baseName = vendorMergeState.customerFile.name.replace(/\.xlsx$/i, '');
    processedFileName = `${baseName}_merged.xlsx`;
    
    resultLink.innerHTML = `
      <button id="select-output-btn" style="
        background: #1f6feb;
        color: white;
        border: none;
        padding: 12px 24px;
        border-radius: 4px;
        cursor: pointer;
        font-weight: 600;
        font-size: 16px;
        width: 100%;
      ">Select Output File</button>
    `;
    resultPanel.classList.add('show');
    showStatus(`✓ Vendor data merged successfully. Choose where to save the Excel file.`, 'success');
    
    setupOutputButton();
    
  } catch (err) {
    console.error('Error:', err);
    showStatus(`Error: ${err.message}`, 'error');
  } finally {
    customerBtn.disabled = false;
    vendorBtn.disabled = false;
    vendorMergeState = { customerFile: null, vendorFile: null };
  }
});

/**
 * Handle input folder selection - filter for PDFs and process CLIENT-SIDE
 */
fileInput.addEventListener('change', async () => {
  if (!fileInput.files || fileInput.files.length === 0) {
    return;
  }
  
  // Filter for PDF files only
  const files = Array.from(fileInput.files).filter(file => 
    file.type.includes('pdf') || file.name.endsWith('.pdf') || file.name.endsWith('.PDF')
  );
  
  if (files.length === 0) {
    showStatus('No PDF files found in the selected folder.', 'error');
    fileInput.value = '';
    return;
  }
  
  showStatus(`Processing ${files.length} PDF file(s)... Please wait.`, 'loading');
  resultPanel.classList.remove('show');
  customerBtn.disabled = true;
  vendorBtn.disabled = true;
  
  try {
    // Use client-side converter (no server)
    const workbook = await window.POConverter.convertMultiplePdfsToExcel(files);
    
    // Convert workbook to blob
    const buffer = await workbook.xlsx.writeBuffer();
    processedBlob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    processedFileName = 'merged-output.xlsx';
    
    resultLink.innerHTML = `
      <button id="select-output-btn" style="
        background: #1f6feb;
        color: white;
        border: none;
        padding: 12px 24px;
        border-radius: 4px;
        cursor: pointer;
        font-weight: 600;
        font-size: 16px;
        width: 100%;
      ">Select Output File</button>
    `;
    resultPanel.classList.add('show');
    showStatus(`✓ Processed ${files.length} PDF file(s). Choose where to save the Excel file.`, 'success');
    
    setupOutputButton();
    
  } catch (err) {
    console.error('Error:', err);
    showStatus(`Error: ${err.message}`, 'error');
  } finally {
    customerBtn.disabled = false;
    vendorBtn.disabled = false;
    fileInput.value = '';
  }
});

/**
 * Setup the output file save button with Save As dialog
 */
function setupOutputButton() {
  const selectBtn = document.getElementById('select-output-btn');
  selectBtn.addEventListener('click', async () => {
    const desiredName = processedFileName || 'merged-output.xlsx';
    try {
      // Try File System Access API for Save As dialog (Chrome/Edge)
      if (window.showSaveFilePicker) {
        const handle = await window.showSaveFilePicker({
          suggestedName: desiredName,
          types: [
            {
              description: 'Excel Workbook',
              accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] }
            }
          ]
        });
        const writable = await handle.createWritable();
        await writable.write(processedBlob);
        await writable.close();
        showStatus(`✓ Saved as ${handle.name}`, 'success');
        clearOutputButton();
        return;
      }
      
      // Legacy IE support
      if (navigator.msSaveOrOpenBlob) {
        navigator.msSaveOrOpenBlob(processedBlob, desiredName);
        showStatus(`✓ Saved as ${desiredName}`, 'success');
        clearOutputButton();
        return;
      }

      // Fallback: auto-download
      const url = URL.createObjectURL(processedBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = desiredName;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      showStatus(`✓ Download started as ${desiredName}`, 'success');
      clearOutputButton();
    } catch (err) {
      if (err.name === 'AbortError') {
        // User cancelled - not an error
        return;
      }
      console.error('Save error:', err);
      showStatus(`Error saving file: ${err.message || err}`, 'error');
    }
  });
}
