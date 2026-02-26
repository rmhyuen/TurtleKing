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

const PO_LIST_CONFIG = {
  fileName: 'PO#List.xlsx',
  startRow: 10,
  startColumn: 2, // Column B
  endColumn: 23   // Column W
};

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
function showStatus(message, type = 'info', useHTML = false) {
  if (useHTML) {
    statusEl.innerHTML = message;
  } else {
    statusEl.textContent = message;
  }
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
  let poLineRecords = [];
  const fileCount = files.length;
  
  try {
    // Use client-side converter (no server)
    let workbook;
    if (isCombined) {
      const conversionResult = await window.POConverter.convertMultiplePdfsToExcel(files, { includeRecords: true });
      workbook = conversionResult.workbook;
      poLineRecords = conversionResult.records || [];
    } else {
      workbook = await window.POConverter.convertMultiplePdfsToExcel(files);
    }
    
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

        showStatus('✓ Merged data uploaded. Appending rows to PO#List.xlsx...', 'loading');
        let poListResult;
        try {
          poListResult = await appendRowsToPoList(poLineRecords, vendorBuffer, vendorData.name, currentTimestamp);
          console.log(`PO#List append complete. Rows appended: ${poListResult.appendedRows}, skipped (duplicates): ${poListResult.skippedRows}`);
        } catch (poListErr) {
          console.error('PO#List append error:', poListErr);
          showStatus(`Error updating PO#List.xlsx: ${poListErr.message}`, 'error');
          return;
        }
        
        let poListMsg = `✓ Complete — ${poListResult.appendedRows + poListResult.skippedRows} row(s) in batch, ${poListResult.appendedRows} added to PO#List.xlsx`;
        let useHTML = false;
        if (poListResult.skippedRows > 0) {
          poListMsg += `, ${poListResult.skippedRows} skipped (already processed)`;
          if (poListResult.skippedSourceList && poListResult.skippedSourceList.length > 0) {
            const srcItems = poListResult.skippedSourceList.map(s => s.replace(/</g, '&lt;').replace(/>/g, '&gt;'));
            poListMsg += `:<br>— ${srcItems.join('<br>— ')}`;
            useHTML = true;
          }
        }
        poListMsg += '.';
        showStatus(poListMsg, 'success', useHTML);
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
 * Fetch PO#List workbook from Google Drive
 */
async function fetchPoListWorkbookFromDrive() {
  try {
    const response = await fetch(`${CONFIG.scriptUrl}?action=download-po-list&apiKey=${CONFIG.apiKey}`);

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    const result = await response.json();
    if (!result.success) {
      throw new Error(result.error || 'Failed to fetch PO#List workbook');
    }

    const data = result.data;
    const binaryString = atob(data.data);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }

    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return new File([blob], data.fileName || PO_LIST_CONFIG.fileName, { type: blob.type });
  } catch (err) {
    throw new Error(`Could not fetch PO#List.xlsx from CustomerData/Beals/Processed. ${err.message}`);
  }
}

/**
 * Upload updated PO#List workbook to Google Drive
 */
async function uploadPoListWorkbookToDrive(blob) {
  try {
    const base64Data = await blobToBase64(blob);

    return await new Promise((resolve, reject) => {
      const xhr = new XMLHttpRequest();

      xhr.onload = function() {
        if (xhr.status >= 200 && xhr.status < 300) {
          const response = JSON.parse(xhr.responseText || '{}');
          if (response.success) {
            resolve(true);
          } else {
            reject(new Error(response.error || 'PO#List upload failed'));
          }
        } else {
          reject(new Error(`Upload failed: ${xhr.status}`));
        }
      };

      xhr.onerror = function() {
        reject(new Error('PO#List upload request failed'));
      };

      xhr.open('POST', `${CONFIG.scriptUrl}?apiKey=${CONFIG.apiKey}`, true);

      const payload = new FormData();
      payload.append('action', 'upload-po-list');
      payload.append('data', base64Data);

      xhr.send(payload);
    });
  } catch (err) {
    throw new Error(`Could not upload updated PO#List.xlsx. ${err.message}`);
  }
}

/**
 * Append combined workflow rows to PO#List workbook
 */
async function appendRowsToPoList(poLineRecords, vendorArrayBuffer, vendorFileName, appRunTimestamp) {
  if (!Array.isArray(poLineRecords) || poLineRecords.length === 0) {
    throw new Error('No PO line records were generated for append');
  }

  const vendorLookup = await buildVendorLookup(vendorArrayBuffer, vendorFileName);
  const appendRows = buildPoListRows(poLineRecords, vendorLookup, appRunTimestamp);

  if (appendRows.length === 0) {
    return { appendedRows: 0, skippedRows: 0 };
  }

  const poListFile = await fetchPoListWorkbookFromDrive();
  const result = await appendRowsToWorkbook(poListFile, appendRows);

  await uploadPoListWorkbookToDrive(result.blob);

  return { appendedRows: result.appendedRows, skippedRows: result.skippedRows, skippedPOList: result.skippedPOList || [], skippedSourceList: result.skippedSourceList || [] };
}

/**
 * Build vendor lookup keyed by Item#/MFG STYLE
 */
async function buildVendorLookup(vendorArrayBuffer, vendorFileName) {
  const vendorLookup = {};
  const isCSV = (vendorFileName || '').toLowerCase().endsWith('.csv');

  if (isCSV) {
    const content = new TextDecoder().decode(vendorArrayBuffer);
    const { rows } = window.POConverter.parseCSV(content);

    for (const row of rows) {
      const itemNum = (row['Item#'] || '').toString().trim();
      const whsNum = parseInt(row['WHS#'], 10) || 0;

      if (itemNum && !vendorLookup[itemNum] && whsNum === 1) {
        const baseCost = parseCurrencyToNumber(row['Base Cost']);
        vendorLookup[itemNum] = {
          vend: (row['Vend#'] || '').toString().trim(),
          baseCost,
          boxCase: (row['Box/Case'] || '').toString().trim(),
          unitCase: (row['Unit/Case'] || '').toString().trim()
        };
      }
    }
  } else {
    const workbook = new window.ExcelJS.Workbook();
    await workbook.xlsx.load(vendorArrayBuffer);

    const worksheet = workbook.worksheets[0];
    if (!worksheet) return vendorLookup;

    const headerRow = worksheet.getRow(1);
    const cols = {};
    headerRow.eachCell((cell, num) => {
      const val = (cell.value || '').toString().trim();
      if (val === 'Item#') cols.item = num;
      else if (val === 'WHS#') cols.whs = num;
      else if (val === 'Vend#') cols.vend = num;
      else if (val === 'Base Cost') cols.baseCost = num;
      else if (val === 'Unit/Case') cols.unitCase = num;
      else if (val === 'Box/Case') cols.boxCase = num;
    });

    if (!cols.item || !cols.whs) {
      return vendorLookup;
    }

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;

      const itemNum = (row.getCell(cols.item).value || '').toString().trim();
      const whsNum = parseInt(row.getCell(cols.whs).value, 10) || 0;

      if (itemNum && !vendorLookup[itemNum] && whsNum === 1) {
        vendorLookup[itemNum] = {
          vend: (row.getCell(cols.vend).value || '').toString().trim(),
          baseCost: parseCurrencyToNumber(row.getCell(cols.baseCost).value),
          boxCase: cols.boxCase ? (row.getCell(cols.boxCase).value || '').toString().trim() : '',
          unitCase: cols.unitCase ? (row.getCell(cols.unitCase).value || '').toString().trim() : ''
        };
      }
    });
  }

  return vendorLookup;
}

/**
 * Map parsed records to PO#List column order B:W
 */
function buildPoListRows(records, vendorLookup, appRunTimestamp) {
  const rows = [];

  for (const record of records) {
    const poNumber = (record['PO #'] || '').toString().trim();
    const sku = (record['SKU'] || '').toString().trim();
    const mfgStyle = (record['MFG Style'] || '').toString().trim();

    if (!poNumber && !sku && !mfgStyle) {
      continue;
    }

    const ttlUnits = parseInt(record['Qty'], 10) || 0;
    const costUnit = parseCurrencyToNumber(record['Cost/Unit']);
    const retail = parseCurrencyToNumber(record['Retail']);
    const packQty = parseInt(record['Pack Qty.'], 10) || '';
    const ttlAmt = costUnit !== null ? costUnit * ttlUnits : '';

    const vendorData = vendorLookup[mfgStyle] || {};
    const baseCost = typeof vendorData.baseCost === 'number' ? vendorData.baseCost : null;
    const ttlCost = baseCost !== null ? baseCost * ttlUnits : '';

    rows.push([
      poNumber,
      (record['ORDER DATE'] || '').toString().trim(),
      (record['SHIP DATE'] || '').toString().trim(),
      (record['CANCEL DATE'] || '').toString().trim(),
      record['DEPT #'] || '',
      (record['DC'] || '').toString().trim(),
      record['STORE #'] || '',
      sku,
      mfgStyle,
      costUnit !== null ? costUnit : '',
      retail !== null ? retail : '',
      packQty,
      ttlUnits,
      ttlAmt,
      record['Line'] || '',
      appRunTimestamp,
      vendorData.vend || '',
      baseCost !== null ? baseCost : '',
      vendorData.boxCase || '',
      vendorData.unitCase || '',
      ttlCost,
      (record['SourceFile'] || '').toString().trim()
    ]);
  }

  return rows;
}

/**
 * Append rows to first worksheet in PO#List workbook
 */
async function appendRowsToWorkbook(poListFile, rowsToAppend) {
  const workbook = new window.ExcelJS.Workbook();
  const arrayBuffer = await poListFile.arrayBuffer();
  await workbook.xlsx.load(arrayBuffer);

  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error('PO#List workbook has no worksheet');
  }

  let nextRow = PO_LIST_CONFIG.startRow;
  const maxRow = Math.max(worksheet.rowCount, PO_LIST_CONFIG.startRow);
  for (let rowNum = PO_LIST_CONFIG.startRow; rowNum <= maxRow; rowNum++) {
    let hasData = false;
    for (let col = PO_LIST_CONFIG.startColumn; col <= PO_LIST_CONFIG.endColumn; col++) {
      const cellValue = worksheet.getCell(rowNum, col).value;
      if (cellValue !== null && cellValue !== undefined && String(cellValue).trim() !== '') {
        hasData = true;
        break;
      }
    }
    if (!hasData) {
      nextRow = rowNum;
      break;
    }
    nextRow = rowNum + 1;
  }

  // --- Duplicate check: skip rows whose SourceFile already exists in PO#List ---
  // SourceFile is in column W (23)
  const srcColIndex = PO_LIST_CONFIG.endColumn;         // Column W (23)

  const existingSources = new Set();
  for (let rowNum = PO_LIST_CONFIG.startRow; rowNum < nextRow; rowNum++) {
    const srcVal = (worksheet.getCell(rowNum, srcColIndex).value ?? '').toString().trim();
    if (srcVal) {
      existingSources.add(srcVal);
    }
  }

  const originalCount = rowsToAppend.length;
  const skippedPOs = new Set();
  const skippedSources = new Set();
  rowsToAppend = rowsToAppend.filter(row => {
    const srcVal = (row[21] ?? '').toString().trim();      // SourceFile
    if (existingSources.has(srcVal)) {
      const poVal = (row[0] ?? '').toString().trim();      // PO#
      if (poVal) skippedPOs.add(poVal);
      if (srcVal) skippedSources.add(srcVal);
      return false;
    }
    return true;
  });

  const skippedCount = originalCount - rowsToAppend.length;
  const skippedPOList = Array.from(skippedPOs).sort();
  const skippedSourceList = Array.from(skippedSources).sort();
  if (skippedCount > 0) {
    console.log(`Duplicate check: skipped ${skippedCount} row(s) already in PO#List.xlsx. POs skipped: ${skippedPOList.join(', ')}. Sources: ${skippedSourceList.join(', ')}`);
  }

  if (rowsToAppend.length === 0) {
    console.log('All rows already exist in PO#List.xlsx — nothing to append.');
    const buffer = await workbook.xlsx.writeBuffer();
    return {
      blob: new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      }),
      appendedRows: 0,
      skippedRows: skippedCount,
      skippedPOList,
      skippedSourceList
    };
  }
  // --- End duplicate check ---

  for (const rowValues of rowsToAppend) {
    rowValues.forEach((value, idx) => {
      worksheet.getCell(nextRow, PO_LIST_CONFIG.startColumn + idx).value = value;
    });

    // Currency formats for K, L, O, S, V
    worksheet.getCell(nextRow, 11).numFmt = '$#,##0.00';
    worksheet.getCell(nextRow, 12).numFmt = '$#,##0.00';
    worksheet.getCell(nextRow, 15).numFmt = '$#,##0.00';
    worksheet.getCell(nextRow, 19).numFmt = '$#,##0.00';
    worksheet.getCell(nextRow, 22).numFmt = '$#,##0.00';

    nextRow++;
  }

  // Sort all populated data rows (B:W) by PO# in column B
  const dataRows = [];
  const lastAppendedRow = Math.max(PO_LIST_CONFIG.startRow, nextRow - 1);
  for (let rowNum = PO_LIST_CONFIG.startRow; rowNum <= lastAppendedRow; rowNum++) {
    const rowValues = [];
    let hasData = false;
    for (let col = PO_LIST_CONFIG.startColumn; col <= PO_LIST_CONFIG.endColumn; col++) {
      const cellValue = worksheet.getCell(rowNum, col).value;
      rowValues.push(cellValue);
      if (cellValue !== null && cellValue !== undefined && String(cellValue).trim() !== '') {
        hasData = true;
      }
    }
    if (hasData) {
      dataRows.push(rowValues);
    }
  }

  dataRows.sort((a, b) => {
    const poA = (a[0] ?? '').toString().trim();
    const poB = (b[0] ?? '').toString().trim();

    const numA = parseInt(poA, 10);
    const numB = parseInt(poB, 10);
    const aIsNum = Number.isFinite(numA);
    const bIsNum = Number.isFinite(numB);

    if (aIsNum && bIsNum) {
      return numA - numB;
    }
    if (aIsNum && !bIsNum) return -1;
    if (!aIsNum && bIsNum) return 1;
    return poA.localeCompare(poB);
  });

  // Clear existing B:W data region, then rewrite sorted rows from row 10
  for (let rowNum = PO_LIST_CONFIG.startRow; rowNum <= lastAppendedRow; rowNum++) {
    for (let col = PO_LIST_CONFIG.startColumn; col <= PO_LIST_CONFIG.endColumn; col++) {
      worksheet.getCell(rowNum, col).value = null;
    }
  }

  let writeRow = PO_LIST_CONFIG.startRow;
  for (const rowValues of dataRows) {
    rowValues.forEach((value, idx) => {
      worksheet.getCell(writeRow, PO_LIST_CONFIG.startColumn + idx).value = value;
    });

    // Currency formats for K, L, O, S, V
    worksheet.getCell(writeRow, 11).numFmt = '$#,##0.00';
    worksheet.getCell(writeRow, 12).numFmt = '$#,##0.00';
    worksheet.getCell(writeRow, 15).numFmt = '$#,##0.00';
    worksheet.getCell(writeRow, 19).numFmt = '$#,##0.00';
    worksheet.getCell(writeRow, 22).numFmt = '$#,##0.00';

    writeRow++;
  }

  const lastDataRow = Math.max(PO_LIST_CONFIG.startRow, writeRow - 1);

  // Row 8 summary formulas for N (TTL UNITS), O (TTL AMT), and V (TTL COST)
  worksheet.getCell(8, 14).value = { formula: `SUM(N${PO_LIST_CONFIG.startRow}:N${lastDataRow})` };
  worksheet.getCell(8, 15).value = { formula: `SUM(O${PO_LIST_CONFIG.startRow}:O${lastDataRow})` };
  worksheet.getCell(8, 22).value = { formula: `SUM(V${PO_LIST_CONFIG.startRow}:V${lastDataRow})` };
  worksheet.getCell(8, 15).numFmt = '$#,##0.00';
  worksheet.getCell(8, 22).numFmt = '$#,##0.00';

  const buffer = await workbook.xlsx.writeBuffer();
  return {
    blob: new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }),
    appendedRows: rowsToAppend.length,
    skippedRows: skippedCount,
    skippedPOList,
    skippedSourceList
  };
}

function parseCurrencyToNumber(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number') return value;
  const parsed = parseFloat(String(value).replace(/[$,]/g, ''));
  return Number.isFinite(parsed) ? parsed : null;
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

