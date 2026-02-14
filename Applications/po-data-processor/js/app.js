/**
 * PO Data Processor - Client-Side Application
 * 100% browser-based - no server required, works on GitHub Pages
 */

const customerBtn = document.getElementById('customer-btn');
const vendorBtn = document.getElementById('vendor-btn');
const combinedBtn = document.getElementById('combined-btn');
const settingsBtn = document.getElementById('settings-btn');
const fileInput = document.getElementById('file-input');
const customerExcelInput = document.getElementById('customer-excel-input');
const vendorDataInput = document.getElementById('vendor-data-input');
const statusEl = document.getElementById('status');
const resultPanel = document.getElementById('result-panel');
const resultLink = document.getElementById('result-link');

let processedBlob = null;
let processedFileName = 'merged-output.xlsx';
let currentTimestamp = '';

/**
 * Generate timestamp in YYYYMMDD_HHMMSS format
 */
function generateTimestamp() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  const seconds = String(now.getSeconds()).padStart(2, '0');
  return `${year}${month}${day}_${hours}${minutes}${seconds}`;
}

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
if (customerBtn) {
  customerBtn.addEventListener('click', () => {
    window.isCombinedWorkflow = false;
    showStatus('Select folder with PDF files.', 'info');
    fileInput.click();
  });
} else {
  console.error('Customer button not found in DOM');
}

/**
 * Handle settings button click
 */
if (settingsBtn) {
  settingsBtn.addEventListener('click', () => {
    SETUP.showSettings();
  });
} else {
  console.error('Settings button not found in DOM');
}

/**
 * Handle combined button click - customer data + vendor merge in one workflow
 */
if (combinedBtn) {
  combinedBtn.addEventListener('click', () => {
    window.isCombinedWorkflow = true;
    showStatus('Select folder with PDF files for customer data.', 'info');
    fileInput.click();
  });
} else {
  console.error('Combined button not found in DOM');
}

/**
 * Handle vendor data button click - fetch customer and vendor data automatically
 */
if (vendorBtn) {
  vendorBtn.addEventListener('click', async () => {
    showStatus('Fetching customer data from Google Drive...', 'loading');
    resultPanel.classList.remove('show');
    customerBtn.disabled = true;
    vendorBtn.disabled = true;
    combinedBtn.disabled = true;
    
    try {
      // Fetch both customer and vendor data from Google Drive
      const customerData = await fetchCustomerDataFromDrive();
      if (!customerData) {
        showStatus('Error: Could not fetch customer data from Google Drive', 'error');
        return;
      }
      
      showStatus('Fetching vendor data from Google Drive...', 'loading');
      
      const vendorData = await fetchVendorDataFromDrive();
      if (!vendorData) {
        showStatus('Error: Could not fetch vendor data from Google Drive', 'error');
        return;
      }
      
      vendorMergeState.customerFile = customerData;
      vendorMergeState.vendorFile = vendorData;
      
      showStatus('Merging customer and vendor data...', 'loading');
      
      // Proceed with merge automatically
      await performVendorMerge();
      
    } catch (err) {
      console.error('Error:', err);
      showStatus(`Error: ${err.message}`, 'error');
    } finally {
      customerBtn.disabled = false;
      vendorBtn.disabled = false;
      combinedBtn.disabled = false;
    }
  });
} else {
  console.error('Vendor button not found in DOM');
}

/**
 * Handle vendor data file selection for vendor merge (CLIENT-SIDE)
 */
if (vendorDataInput) {
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
  combinedBtn.disabled = true;
  
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
    
    // Generate output filename with timestamp
    currentTimestamp = generateTimestamp();
    const baseName = vendorMergeState.customerFile.name.replace(/\.xlsx$/i, '');
    processedFileName = `${baseName}_${currentTimestamp}.xlsx`;
    
    resultPanel.classList.remove('show');
    showStatus(`✓ Vendor data merged successfully. Uploading to Google Drive...`, 'success');
    
    setupOutputButton(true);
    
  } catch (err) {
    console.error('Error:', err);
    showStatus(`Error: ${err.message}`, 'error');
  } finally {
    customerBtn.disabled = false;
    vendorBtn.disabled = false;
    combinedBtn.disabled = false;
    vendorMergeState = { customerFile: null, vendorFile: null };
  }
  });
} else {
  console.error('Vendor data input not found in DOM');
}

/**
 * Handle input folder selection - filter for PDFs and process CLIENT-SIDE
 */
if (fileInput) {
  fileInput.addEventListener('change', async () => {
    if (!fileInput.files || fileInput.files.length === 0) {
      window.isCombinedWorkflow = false;
      return;
    }
    
    // Filter for PDF files only
    const files = Array.from(fileInput.files).filter(file => 
      file.type.includes('pdf') || file.name.endsWith('.pdf') || file.name.endsWith('.PDF')
    );
  
  if (files.length === 0) {
    showStatus('No PDF files found in the selected folder.', 'error');
    fileInput.value = '';
    window.isCombinedWorkflow = false;
    return;
  }
  
  // Extract parent folder name from webkitRelativePath
  let parentFolderName = 'CustomerData'; // default fallback
  if (files[0] && files[0].webkitRelativePath) {
    const parts = files[0].webkitRelativePath.split('/');
    if (parts.length > 1) {
      parentFolderName = parts[0]; // Get the first part (parent folder)
    }
  }
  
  console.log('Parent folder name: ' + parentFolderName);
  console.log('Number of files: ' + files.length);
  
  showStatus(`Processing ${files.length} PDF file(s)... Please wait.`, 'loading');
  resultPanel.classList.remove('show');
  customerBtn.disabled = true;
  vendorBtn.disabled = true;
  combinedBtn.disabled = true;
  
  const isCombined = window.isCombinedWorkflow;
  let customerDataBlob = null;
  let customerDataFileName = null;
  const fileCount = files.length;
  
  try {
    // Use client-side converter (no server)
    const workbook = await window.POConverter.convertMultiplePdfsToExcel(files);
    
    // Convert workbook to blob
    const buffer = await workbook.xlsx.writeBuffer();
    customerDataBlob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    
    // Generate filename with parent folder, file count, and timestamp
    currentTimestamp = generateTimestamp();
    customerDataFileName = `${parentFolderName}_${fileCount}_${currentTimestamp}.xlsx`;
    
    console.log('Customer data filename: ' + customerDataFileName);
    
    resultPanel.classList.remove('show');
    
    // If combined workflow, fetch vendor data and merge, but also upload customer data
    if (isCombined) {
      showStatus(`✓ Processed ${fileCount} PDF file(s). Uploading customer data...`, 'loading');
      
      // Upload customer data first
      console.log('Combined workflow: uploading customer data...');
      processedBlob = customerDataBlob;
      processedFileName = customerDataFileName;
      const customerUploadSuccess = await uploadToGoogleDrive(processedBlob, currentTimestamp, false, fileCount, parentFolderName);
      
      if (!customerUploadSuccess) {
        showStatus('Error: Could not upload customer data to Google Drive. Skipping vendor merge.', 'error');
        return;
      }
      
      console.log('Customer data uploaded, now fetching vendor data...');
      showStatus(`✓ Customer data uploaded. Fetching vendor data...`, 'loading');
      
      try {
        const vendorData = await fetchVendorDataFromDrive();
        if (!vendorData) {
          console.error('Vendor data fetch returned null/undefined');
          showStatus('Error: Could not fetch vendor data from Google Drive. Customer data uploaded, but merge skipped.', 'error');
          return;
        }
        
        console.log('Vendor data fetched, merging...');
        showStatus('Merging customer and vendor data...', 'loading');
        
        // Merge the customer data with vendor data
        const customerBuffer = await customerDataBlob.arrayBuffer();
        const vendorBuffer = await vendorData.arrayBuffer();
        
        console.log('Calling mergeVendorData...');
        const mergedWorkbook = await window.POConverter.mergeVendorData(
          customerBuffer,
          vendorBuffer,
          vendorData.name
        );
        
        console.log('Merge successful, creating blob...');
        const mergedBuffer = await mergedWorkbook.xlsx.writeBuffer();
        processedBlob = new Blob([mergedBuffer], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        
        // Set filename for merged data - use first 2 parts (folder name + file count) + timestamp
        processedFileName = `${parentFolderName}_${files.length}_${currentTimestamp}.xlsx`;
        
        console.log('Updated filename for merged: ' + processedFileName);
        showStatus(`✓ Processed and merged data. Uploading merged file...`, 'loading');
        
        // Upload the merged data
        const mergedUploadSuccess = await uploadToGoogleDrive(processedBlob, currentTimestamp, true, fileCount, parentFolderName);
        
        if (!mergedUploadSuccess) {
          showStatus('Error: Could not upload merged data, but customer data was uploaded successfully.', 'error');
          return;
        }
        
        showStatus(`✓ Both customer and vendor data uploaded successfully to Google Drive.`, 'success');
        resultLink.innerHTML = `
          <button id="download-btn" style="
            background: #1f6feb;
            color: white;
            border: none;
            padding: 12px 24px;
            border-radius: 4px;
            cursor: pointer;
            font-weight: 600;
            font-size: 16px;
            width: 100%;
          ">Download Locally (Optional)</button>
        `;
        resultPanel.classList.add('show');
        
        document.getElementById('download-btn').addEventListener('click', async () => {
          await saveLocally(processedFileName);
        });
        
      } catch (err) {
        console.error('Error in vendor merge:', err);
        console.error('Error stack:', err.stack);
        showStatus(`Error during vendor merge: ${err.message}. Customer data was uploaded successfully.`, 'error');
        return;
      }
    } else {
      // Regular customer-only workflow
      processedBlob = customerDataBlob;
      processedFileName = customerDataFileName;
      showStatus(`✓ Processed ${fileCount} PDF file(s). Uploading to Google Drive...`, 'success');
      setupOutputButton(false, fileCount, parentFolderName);
    }
    
  } catch (err) {
    console.error('Error:', err);
    showStatus(`Error: ${err.message}`, 'error');
  } finally {
    customerBtn.disabled = false;
    vendorBtn.disabled = false;
    combinedBtn.disabled = false;
    fileInput.value = '';
    window.isCombinedWorkflow = false;
  }
  });
} else {
  console.error('File input not found in DOM');
}

/**
 * Automatically upload processed file to Google Drive
 */
async function setupOutputButton(isMerged = false, fileCount = 0, parentFolderName = '') {
  try {
    // Upload to Google Drive automatically
    console.log('setupOutputButton() called');
    console.log('Current filename: ' + processedFileName);
    console.log('Current blob size: ' + processedBlob.size);
    console.log('Is merged: ' + isMerged);
    console.log('File count: ' + fileCount);
    console.log('Parent folder name: ' + parentFolderName);
    
    showStatus(`Uploading to Google Drive...`, 'loading');
    
    const uploadSuccess = await uploadToGoogleDrive(processedBlob, currentTimestamp, isMerged, fileCount, parentFolderName);
    
    console.log('Upload result: ' + uploadSuccess);
    
    if (!uploadSuccess) {
      showStatus(`Error: Could not upload to Google Drive. Please try again.`, 'error');
      return;
    }
    
    const desiredName = processedFileName || 'merged-output.xlsx';
    
    // Show success and offer download option
    resultLink.innerHTML = `
      <button id="download-btn" style="
        background: #1f6feb;
        color: white;
        border: none;
        padding: 12px 24px;
        border-radius: 4px;
        cursor: pointer;
        font-weight: 600;
        font-size: 16px;
        width: 100%;
      ">Download Locally (Optional)</button>
    `;
    resultPanel.classList.add('show');
    showStatus(`✓ Uploaded to Google Drive. Optionally download a local copy.`, 'success');
    
    document.getElementById('download-btn').addEventListener('click', async () => {
      await saveLocally(desiredName);
    });
    
  } catch (err) {
    console.error('Error:', err);
    showStatus(`Error: ${err.message || err}`, 'error');
  }
}

/**
 * Fetch customer data from Google Drive (latest from CustomerDataOnly folder)
 */
async function fetchCustomerDataFromDrive() {
  try {
    const response = await fetch(`${CONFIG.scriptUrl}?action=download-customer-data-only&apiKey=${CONFIG.apiKey}`);
    
    if (!response.ok) {
      console.error('HTTP Error:', response.status, response.statusText);
      return null;
    }
    
    const result = await response.json();
    
    if (!result.success) {
      console.error('Error from server:', result.error);
      return null;
    }
    
    const data = result.data;
    
    // Convert base64 to blob and then to File
    const binaryString = atob(data.data);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const file = new File([blob], data.fileName, { type: blob.type });
    
    return file;
  } catch (err) {
    console.error('Fetch customer data error:', err);
    return null;
  }
}

/**
 * Fetch vendor data from Google Drive
 */
async function fetchVendorDataFromDrive() {
  try {
    const response = await fetch(`${CONFIG.scriptUrl}?action=download-vendor-file&apiKey=${CONFIG.apiKey}`);
    
    if (!response.ok) {
      console.error('HTTP Error:', response.status, response.statusText);
      return null;
    }
    
    const result = await response.json();
    
    if (!result.success) {
      console.error('Error from server:', result.error);
      return null;
    }
    
    const data = result.data;
    
    // Convert base64 to blob and then to File
    const binaryString = atob(data.data);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const file = new File([blob], data.fileName, { type: blob.type });
    
    return file;
  } catch (err) {
    console.error('Fetch vendor data error:', err);
    return null;
  }
}

/**
 * Perform vendor merge with downloaded vendor file
 */
async function performVendorMerge() {
  try {
    // Validate we have both files
    if (!vendorMergeState.customerFile || !vendorMergeState.vendorFile) {
      showStatus('Error: Missing files. Please try again.', 'error');
      return;
    }
    
    customerBtn.disabled = true;
    vendorBtn.disabled = true;
    
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
    
    // Generate output filename with timestamp
    
    // Extract parent folder name, file count, and timestamp from customer data filename
    // Expected format: {parentFolder}_{fileCount}_{timestamp}.xlsx or {parentFolder}_{timestamp}.xlsx
    let parentFolderName = 'Beals'; // default fallback
    let fileCount = ''; // default fallback (empty string means no file count)
    let extractedTimestamp = generateTimestamp(); // default fallback
    
    if (vendorMergeState.customerFile && vendorMergeState.customerFile.name) {
      const filename = vendorMergeState.customerFile.name.replace('.xlsx', '');
      const nameParts = filename.split('_');
      
      if (nameParts.length >= 2) {
        parentFolderName = nameParts[0];
        
        if (nameParts.length >= 4) {
          // Format: parentFolder_fileCount_YYYYMMDD_HHMMSS
          fileCount = nameParts[1];
          extractedTimestamp = `${nameParts[2]}_${nameParts[3]}`;
        } else if (nameParts.length === 3) {
          // Format: parentFolder_YYYYMMDD_HHMMSS
          extractedTimestamp = `${nameParts[1]}_${nameParts[2]}`;
          fileCount = ''; // no file count in this format
        } else {
          // Format: parentFolder_timestamp (single part timestamp)
          extractedTimestamp = nameParts[1];
          fileCount = '';
        }
      }
    }
    
    currentTimestamp = extractedTimestamp;
    processedFileName = fileCount ? `${parentFolderName}_${fileCount}_${currentTimestamp}.xlsx` : `${parentFolderName}_${currentTimestamp}.xlsx`;
    
    // Automatically upload to Google Drive using new function with proper parameters
    showStatus('Uploading merged data to Google Drive...', 'loading');
    const uploadSuccess = await uploadToGoogleDrive(processedBlob, currentTimestamp, true, fileCount, parentFolderName);
    
    if (!uploadSuccess) {
      showStatus('Error: Could not upload merged data to Google Drive.', 'error');
      return;
    }
    
    // Show success and download option
    resultLink.innerHTML = `
      <button id="download-btn" style="
        background: #1f6feb;
        color: white;
        border: none;
        padding: 12px 24px;
        border-radius: 4px;
        cursor: pointer;
        font-weight: 600;
        font-size: 16px;
        width: 100%;
      ">Download Locally (Optional)</button>
    `;
    resultPanel.classList.add('show');
    showStatus(`✓ Uploaded merged data to Google Drive. Optionally download a local copy.`, 'success');
    
    document.getElementById('download-btn').addEventListener('click', async () => {
      await saveLocally(processedFileName);
    });
    
  } catch (err) {
    console.error('Error:', err);
    showStatus(`Error: ${err.message}`, 'error');
  } finally {
    customerBtn.disabled = false;
    vendorBtn.disabled = false;
  }
}

/**
 * Upload merged customer and vendor data to Google Drive
 */
// Note: uploadMergedDataToDrive() has been replaced by uploadToGoogleDrive() function
// which supports dynamic filenames with parent folder names and file counts

/**
 * Upload file to Google Drive with dynamic naming
 */
async function uploadToGoogleDrive(blob, timestamp, isMerged, fileCount, parentFolderName = '') {
  try {
    // Get base64 encoded data
    const base64Data = await blobToBase64(blob);
    
    // Use parent folder name or fallback to defaults
    let filePrefix = parentFolderName || (isMerged ? 'CustomerAndVendorData' : 'CustomerData');
    
    console.log('Upload action based on merge status: ' + (isMerged ? 'upload-merged-customer-vendor-data' : 'upload-processed-customer-data'));
    console.log('File prefix: ' + filePrefix);
    console.log('Is merged: ' + isMerged);
    console.log('File count: ' + fileCount);
    console.log('Parent folder name: ' + parentFolderName);
    
    // Call Google Apps Script deployment endpoint via XMLHttpRequest
    return new Promise((resolve, reject) => {
      const xhr = new XMLHttpRequest();
      
      xhr.onload = function() {
        if (xhr.status >= 200 && xhr.status < 300) {
          console.log('Upload successful');
          resolve(true);
        } else {
          console.error('Upload failed with status:', xhr.status);
          reject(new Error(`Upload failed: ${xhr.status}`));
        }
      };
      
      xhr.onerror = function() {
        console.error('Upload error');
        reject(new Error('Upload request failed'));
      };
      
      xhr.open('POST', `${CONFIG.scriptUrl}?apiKey=${CONFIG.apiKey}`, true);
      
      const payload = new FormData();
      payload.append('action', isMerged ? 'upload-merged-customer-vendor-data' : 'upload-processed-customer-data');
      payload.append('fileName', filePrefix);
      payload.append('fileCount', fileCount);
      payload.append('timestamp', timestamp);
      payload.append('data', base64Data);
      
      xhr.send(payload);
    });
    
  } catch (err) {
    console.error('Upload error:', err);
    return false;
  }
}
/**
 * Convert blob to Base64
 */
function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      // Remove data:application/...;base64, prefix
      const base64 = reader.result.split(',')[1];
      resolve(base64);
    };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

/**
 * Save file locally
 */
async function saveLocally(desiredName) {
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
      showStatus(`✓ Saved locally as ${handle.name}`, 'success');
      clearOutputButton();
      return;
    }
    
    // Legacy IE support
    if (navigator.msSaveOrOpenBlob) {
      navigator.msSaveOrOpenBlob(processedBlob, desiredName);
      showStatus(`✓ Saved locally as ${desiredName}`, 'success');
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
}

