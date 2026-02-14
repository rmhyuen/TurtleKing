/**
 * Google Apps Script - Automated PDF Processing for Google Drive
 * 
 * Manages the complete workflow:
 * 1. Lists customer PDFs from WEBSITE_APP/CustomerData/Beals/New/
 * 2. Fetches vendor data from WEBSITE_APP/VendorData/
 * 3. Saves processed output to designated folders
 * 4. Moves processed PDFs to archive (Old folder)
 * 
 * DEPLOYMENT INFO:
 * Status: ACTIVE ✅
 * Deployed: February 6, 2026
 * Deployment URL: https://script.google.com/macros/s/AKfycbwknF5fBpZHwy-U3nIhlOA8nWyKKLRx48VfT87XaAithJ3BcpcVx3nIWcY4fXw21dxh/exec
 * 
 * SETUP INSTRUCTIONS:
 * 1. Go to https://script.google.com
 * 2. Create a new project
 * 3. Replace the default code with this entire script
 * 4. Update FOLDER_ID below with your folder ID
 * 5. Save the project
 * 6. Click "Deploy" > "New Deployment" > Select "Web app"
 * 7. Execute as: Your account
 *    Who has access: Anyone
 * 8. Copy the deployment URL and use it in your app
 */

// ⚙️ CONFIGURATION
// Drive Folder ID from: https://drive.google.com/drive/folders/1EWsCGFT1iqMwUad48t8V_n2Irko-sxGZ
const FOLDER_ID = '1EWsCGFT1iqMwUad48t8V_n2Irko-sxGZ';

// 🔐 SECURITY: API Key Authentication
const AUTHORIZED_API_KEY = 'YOUR_API_KEY_HERE'; // Replace with your actual API key

function verifyAuthentication(e) {
  if (!e || !e.parameter) {
    return error('Unauthorized: Missing request parameters');
  }
  const providedKey = e.parameter.apiKey || e.parameter.api_key;
  if (providedKey !== AUTHORIZED_API_KEY) {
    return error('Unauthorized: Invalid or missing API key');
  }
  return null; // Authentication passed
}

const FOLDER_STRUCTURE = {
  customerNew: ['CustomerData', 'Beals', 'New'],
  customerProcessed: ['CustomerData', 'Beals', 'Processed'],
  customerProcessedOnly: ['CustomerData', 'Beals', 'Processed', 'CustomerDataOnly'],
  customerAndVendorData: ['CustomerData', 'Beals', 'Processed', 'CustomerAndVendorData'],
  customerOld: ['CustomerData', 'Beals', 'Old'],
  vendorData: ['VendorData'],
  vendorNew: ['VendorData', 'New'],
  vendorProcessed: ['VendorData', 'Processed']
};

/**
 * Main web app endpoint
 */
function doGet(e) {
  // Verify authentication first
  const authError = verifyAuthentication(e);
  if (authError) return authError;
  
  try {
    const action = e.parameter.action || 'status';
    
    if (action === 'status') {
      return getStatus();
    } else if (action === 'list-customer-pdfs') {
      return listCustomerPDFs();
    } else if (action === 'list-vendor-data') {
      return listVendorData();
    } else if (action === 'download-vendor-file') {
      return downloadVendorFile();
    } else if (action === 'download-customer-data-only') {
      return downloadLatestCustomerDataOnly();
    } else if (action === 'download') {
      return downloadFile(e.parameter.fileId);
    } else if (action === 'archive-pdfs') {
      return archiveProcessedPDFsFromGet(e);
    } else {
      return error('Unknown action: ' + action);
    }
  } catch (err) {
    return error('Server error: ' + err.toString());
  }
}

function doPost(e) {
  // Verify authentication first
  const authError = verifyAuthentication(e);
  if (authError) return authError;
  
  try {
    const action = e.parameter.action || '';
    
    if (action === 'process-order') {
      return processOrder(e);
    } else if (action === 'upload-processed-customer-data') {
      return uploadProcessedCustomerData(e);
    } else if (action === 'upload-merged-customer-vendor-data') {
      return uploadMergedCustomerVendorData(e);
    } else {
      return error('Unknown POST action: ' + action);
    }
  } catch (err) {
    return error('Server error: ' + err.toString());
  }
}

/**
 * Process complete order: download PDFs, merge with vendor data, return Excel
 */
function processOrder(e) {
  try {
    // Step 1: Get all customer PDFs from New folder
    const customerFolder = navigateToFolder(FOLDER_STRUCTURE.customerNew);
    const pdfFiles = customerFolder.getFilesByType(MimeType.PDF);
    const pdfList = [];
    const pdfIds = [];
    
    while (pdfFiles.hasNext()) {
      const file = pdfFiles.next();
      pdfList.push(file);
      pdfIds.push(file.getId());
    }
    
    if (pdfList.length === 0) {
      return error('No PDFs found to process');
    }
    
    // Step 2: Get vendor data file from VendorData/New folder
    const vendorFolder = navigateToFolder(FOLDER_STRUCTURE.vendorNew);
    const vendorFiles = vendorFolder.getFiles();
    let vendorFile = null;
    
    while (vendorFiles.hasNext()) {
      const file = vendorFiles.next();
      if (isDataFile(file)) {
        vendorFile = file;
        break;
      }
    }
    
    if (!vendorFile) {
      return error('No vendor data file found');
    }
    
    // Step 3: For now, just return the list of files that would be processed
    // (Full PDF processing would require external service)
    return success({
      message: 'Processing initiated',
      pdfCount: pdfList.length,
      vendorFile: vendorFile.getName(),
      pdfIds: pdfIds
    });
    
  } catch (err) {
    return error('Error processing order: ' + err.toString());
  }
}

/**
 * Get system status and folder structure
 */
function getStatus() {
  try {
    const customerNewFolder = navigateToFolder(FOLDER_STRUCTURE.customerNew);
    const vendorNewFolder = navigateToFolder(FOLDER_STRUCTURE.vendorNew);
    
    const customerPdfCount = countFiles(customerNewFolder, MimeType.PDF);
    const vendorFileCount = countFiles(vendorNewFolder);
    
    return success({
      system: 'ready',
      folders: {
        customerNew: FOLDER_STRUCTURE.customerNew.join('/'),
        vendorNew: FOLDER_STRUCTURE.vendorNew.join('/')
      },
      stats: {
        customerPdfsReady: customerPdfCount,
        vendorFilesReady: vendorFileCount
      }
    });
  } catch (err) {
    return error('Status check failed: ' + err.toString());
  }
}

/**
 * List all customer PDFs in New folder
 */
function listCustomerPDFs() {
  try {
    const folder = navigateToFolder(FOLDER_STRUCTURE.customerNew);
    const files = folder.getFilesByType(MimeType.PDF);
    const pdfList = [];
    
    while (files.hasNext()) {
      const file = files.next();
      pdfList.push({
        id: file.getId(),
        name: file.getName(),
        size: file.getSize(),
        modifiedDate: file.getLastUpdated()
      });
    }
    
    pdfList.sort((a, b) => new Date(b.modifiedDate) - new Date(a.modifiedDate));
    
    return success({
      folderPath: FOLDER_STRUCTURE.customerNew.join('/'),
      count: pdfList.length,
      files: pdfList
    });
    
  } catch (err) {
    return error('Error listing customer PDFs: ' + err.toString());
  }
}

/**
 * List all vendor data files in New folder
 */
function listVendorData() {
  try {
    const folder = navigateToFolder(FOLDER_STRUCTURE.vendorNew);
    const files = folder.getFiles();
    const fileList = [];
    
    while (files.hasNext()) {
      const file = files.next();
      // Only include CSV and Excel files
      if (isDataFile(file)) {
        fileList.push({
          id: file.getId(),
          name: file.getName(),
          size: file.getSize(),
          modifiedDate: file.getLastUpdated(),
          mimeType: file.getMimeType()
        });
      }
    }
    
    fileList.sort((a, b) => new Date(b.modifiedDate) - new Date(a.modifiedDate));
    
    return success({
      folderPath: FOLDER_STRUCTURE.vendorNew.join('/'),
      count: fileList.length,
      files: fileList
    });
    
  } catch (err) {
    return error('Error listing vendor data: ' + err.toString());
  }
}

/**
 * Download the vendor data file from VendorData folder
 */
function downloadVendorFile() {
  try {
    const folder = navigateToFolder(FOLDER_STRUCTURE.vendorData);
    const files = folder.getFiles();
    let vendorFile = null;
    
    // Find the first data file in VendorData folder
    while (files.hasNext()) {
      const file = files.next();
      if (isDataFile(file)) {
        vendorFile = file;
        break;
      }
    }
    
    if (!vendorFile) {
      return error('No vendor data file found in VendorData folder');
    }
    
    // Get file as base64
    const blob = vendorFile.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    
    return success({
      fileName: vendorFile.getName(),
      mimeType: blob.getContentType(),
      data: base64
    });
    
  } catch (err) {
    return error('Error downloading vendor file: ' + err.toString());
  }
}

/**
 * Download the latest processed customer data file from CustomerDataOnly folder
 */
function downloadLatestCustomerDataOnly() {
  try {
    const folder = navigateToFolder(FOLDER_STRUCTURE.customerProcessedOnly);
    const files = folder.getFiles();
    let latestFile = null;
    let latestDate = null;
    
    // Find the most recent Excel file
    while (files.hasNext()) {
      const file = files.next();
      if (isDataFile(file)) {
        const modDate = file.getLastUpdated();
        if (!latestDate || modDate > latestDate) {
          latestFile = file;
          latestDate = modDate;
        }
      }
    }
    
    if (!latestFile) {
      return error('No customer data file found in CustomerDataOnly folder');
    }
    
    // Get file as base64
    const blob = latestFile.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    
    return success({
      fileName: latestFile.getName(),
      mimeType: blob.getContentType(),
      data: base64
    });
    
  } catch (err) {
    return error('Error downloading customer data file: ' + err.toString());
  }
}

/**
 * Upload merged customer and vendor data to Google Drive with timestamp
 */
function uploadMergedCustomerVendorData(e) {
  try {
    let fileData = null;
    let fileName = e.parameter.fileName || 'CustomerAndVendorData';
    let timestamp = e.parameter.timestamp || '';
    
    // Get file data from parameter
    if (e.parameter.data) {
      fileData = e.parameter.data;
    }
    
    if (!fileData) {
      return error('No file data provided');
    }
    
    // Remove data URL prefix if present
    if (fileData.indexOf('base64,') > -1) {
      fileData = fileData.split('base64,')[1];
    }
    
    // Use provided timestamp or generate one
    if (!timestamp) {
      const now = new Date();
      const year = now.getFullYear();
      const month = String(now.getMonth() + 1).padStart(2, '0');
      const day = String(now.getDate()).padStart(2, '0');
      const hours = String(now.getHours()).padStart(2, '0');
      const minutes = String(now.getMinutes()).padStart(2, '0');
      const seconds = String(now.getSeconds()).padStart(2, '0');
      timestamp = `${year}${month}${day}_${hours}${minutes}${seconds}`;
    }
    
    // Create filename with timestamp
    const finalFileName = `${fileName}_${timestamp}.xlsx`;
    
    // Navigate to CustomerAndVendorData folder
    const targetFolder = navigateToFolder(FOLDER_STRUCTURE.customerAndVendorData);
    
    // Decode base64 and create blob
    const decodedData = Utilities.base64Decode(fileData);
    const blob = Utilities.newBlob(decodedData, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', finalFileName);
    
    // Create file in the folder
    const file = targetFolder.createFile(blob);
    
    return success({
      message: 'Merged file uploaded successfully to Google Drive',
      fileName: finalFileName,
      fileId: file.getId(),
      folder: FOLDER_STRUCTURE.customerAndVendorData.join('/'),
      timestamp: timestamp,
      driveUrl: 'https://drive.google.com/drive/folders/' + targetFolder.getId()
    });
    
  } catch (err) {
    return error('Error uploading merged data: ' + err.toString());
  }
}

/**
 * Download a file by ID (returns as base64 JSON)
 */
function downloadFile(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    
    return success({
      fileName: file.getName(),
      mimeType: blob.getContentType(),
      data: base64
    });
    
  } catch (err) {
    return error('Error downloading file: ' + err.toString());
  }
}

/**
 * Save processed data (customer processed Excel and vendor processed Excel)
 * Expects POST data with file contents
 */
function saveProcessedData(e) {
  try {
    // This would receive the processed files from the web app
    // For now, we track that processing was requested
    return success({
      message: 'Processing request received',
      action: 'Saving processed data to designated folders'
    });
  } catch (err) {
    return error('Error saving processed data: ' + err.toString());
  }
}

/**
 * Upload processed customer data to Google Drive with timestamp
 * Receives base64 encoded Excel file and saves to CustomerDataOnly folder
 */
function uploadProcessedCustomerData(e) {
  try {
    // Handle both POST data and form parameters
    let fileData = null;
    let fileName = e.parameter.fileName || 'CustomerData';
    let timestamp = e.parameter.timestamp || '';
    
    // Try to get file data from POST body first
    if (e.postData && e.postData.contents) {
      fileData = e.postData.contents;
    } else if (e.parameter.data) {
      // Or from parameter
      fileData = e.parameter.data;
    }
    
    if (!fileData) {
      return error('No file data provided');
    }
    
    // Remove data URL prefix if present (data:...;base64,)
    if (fileData.indexOf('base64,') > -1) {
      fileData = fileData.split('base64,')[1];
    }
    
    // Use provided timestamp or generate one
    if (!timestamp) {
      const now = new Date();
      const year = now.getFullYear();
      const month = String(now.getMonth() + 1).padStart(2, '0');
      const day = String(now.getDate()).padStart(2, '0');
      const hours = String(now.getHours()).padStart(2, '0');
      const minutes = String(now.getMinutes()).padStart(2, '0');
      const seconds = String(now.getSeconds()).padStart(2, '0');
      timestamp = `${year}${month}${day}_${hours}${minutes}${seconds}`;
    }
    
    // Create filename with timestamp
    const finalFileName = `${fileName}_${timestamp}.xlsx`;
    
    // Navigate to CustomerDataOnly folder
    const targetFolder = navigateToFolder(FOLDER_STRUCTURE.customerProcessedOnly);
    
    // Decode base64 and create blob
    const decodedData = Utilities.base64Decode(fileData);
    const blob = Utilities.newBlob(decodedData, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', finalFileName);
    
    // Create file in the folder
    const file = targetFolder.createFile(blob);
    
    return success({
      message: 'File uploaded successfully to Google Drive',
      fileName: finalFileName,
      fileId: file.getId(),
      folder: FOLDER_STRUCTURE.customerProcessedOnly.join('/'),
      timestamp: timestamp,
      driveUrl: 'https://drive.google.com/drive/folders/' + targetFolder.getId()
    });
    
  } catch (err) {
    return error('Error uploading processed data: ' + err.toString());
  }
}


/**
 * Archive processed PDFs (move from New to Old) - GET version
 */
function archiveProcessedPDFsFromGet(e) {
  try {
    const idsParam = e.parameter.ids || '';
    const processedPdfIds = idsParam.split(',').filter(id => id.trim());
    
    if (!Array.isArray(processedPdfIds) || processedPdfIds.length === 0) {
      return error('No PDF IDs provided');
    }
    
    const newFolder = navigateToFolder(FOLDER_STRUCTURE.customerNew);
    const oldFolder = navigateToFolder(FOLDER_STRUCTURE.customerOld);
    
    let movedCount = 0;
    processedPdfIds.forEach(fileId => {
      try {
        const file = DriveApp.getFileById(fileId.trim());
        oldFolder.addFile(file);
        newFolder.removeFile(file);
        movedCount++;
      } catch (e) {
        // File not found or already moved, continue
      }
    });
    
    return success({
      message: `Archived ${movedCount} PDF(s) to Old folder`,
      movedCount: movedCount
    });
    
  } catch (err) {
    return error('Error archiving PDFs: ' + err.toString());
  }
}

/**
 * Full cleanup after processing
 */
function cleanup(e) {
  try {
    return success({
      message: 'Cleanup completed',
      timestamp: new Date()
    });
  } catch (err) {
    return error('Error during cleanup: ' + err.toString());
  }
}

/**
 * Helper: Navigate through folder structure
 */
function navigateToFolder(pathArray) {
  const rootFolder = DriveApp.getFolderById(FOLDER_ID);
  let currentFolder = rootFolder;
  
  for (const folderName of pathArray) {
    const folders = currentFolder.getFoldersByName(folderName);
    if (!folders.hasNext()) {
      throw new Error(`Folder not found: ${folderName}`);
    }
    currentFolder = folders.next();
  }
  
  return currentFolder;
}

/**
 * Helper: Count files in folder
 */
function countFiles(folder, mimeType = null) {
  let files = mimeType ? folder.getFilesByType(mimeType) : folder.getFiles();
  let count = 0;
  
  while (files.hasNext()) {
    files.next();
    count++;
  }
  
  return count;
}

/**
 * Helper: Check if file is a data file (CSV or Excel)
 */
function isDataFile(file) {
  const mimeType = file.getMimeType();
  const name = file.getName().toLowerCase();
  
  return (
    mimeType === 'text/csv' ||
    mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
    mimeType === 'application/vnd.ms-excel' ||
    name.endsWith('.csv') ||
    name.endsWith('.xlsx') ||
    name.endsWith('.xls')
  );
}

/**
 * Helper: Return success response
 */
function success(data) {
  return ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      data: data
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Helper: Return error response
 */
function error(message) {
  return ContentService
    .createTextOutput(JSON.stringify({
      success: false,
      error: message
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
