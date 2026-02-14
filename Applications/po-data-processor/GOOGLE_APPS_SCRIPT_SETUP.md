# Google Apps Script Setup Guide

This guide walks you through deploying the backend Google Apps Script that powers the PO Data Processor.

---

## Quick Navigation

- **New Users**: Start with [README.md](README.md) for feature overview
- **Security Setup**: See [SECURITY_SETUP.md](SECURITY_SETUP.md) (required before deploying)
- **Deployment Info**: See [DEPLOYMENT_INFO.md](DEPLOYMENT_INFO.md)

---

## ✅ DEPLOYMENT INFORMATION

**Status:** Deployed and Active  
**Deployment URL:**
```
https://script.google.com/macros/s/AKfycbwknF5fBpZHwy-U3nIhlOA8nWyKKLRx48VfT87XaAithJ3BcpcVx3nIWcY4fXw21dxh/exec
```

**Google Drive Folder ID:** `1EWsCGFT1iqMwUad48t8V_n2Irko-sxGZ`

**Authentication:** API Key Required ✅ (configured via setup modal on first visit)

---

## Overview

The Apps Script serves as a **secure bridge** between your web app and Google Drive:
- ✅ **Automatic file uploads** with timestamps
- ✅ **Automatic file downloads** for vendor merge
- ✅ **API Key authentication** (no unauthorized access)
- ✅ **Handles PDF processing** workflow
- ✅ **Works on GitHub Pages** (100% client-side + secure backend)

## Step-by-Step Setup

### Step 1: Go to Google Apps Script
1. Open https://script.google.com in your browser
2. Click "New project" (or "Create project")
3. Name it something like "PDF Lister" (doesn't matter much)

### Step 2: Copy the Script Code
1. Delete the default `myFunction()` code
2. Open [GoogleAppsScript_Template.example.gs](GoogleAppsScript_Template.example.gs) from this project folder
3. Copy **the entire contents**
4. Paste it into the Google Apps Script editor

### Step 3: Configure Your API Key
The script has a placeholder API key:
```javascript
const AUTHORIZED_API_KEY = 'YOUR_API_KEY_HERE';
```
- **This is the key users will enter in the setup modal**

### Step 4: Save the Project
- Press **Ctrl+S** (Windows) or **Cmd+S** (Mac)
- Or click the "Save" button

### Step 5: Deploy as Web App
1. Click the blue **"Deploy"** button (top right)
2. Click **"New Deployment"**
3. Click the gear icon ⚙️ to show deployment settings
4. Select **"Web app"** from the dropdown
5. Under "Execute as" → Select **your Google account**
6. Under "Who has access" → Select **"Anyone"**
7. Click **"Deploy"**

### Step 6: Deployment Complete
1. You'll see a popup with your deployment URL
2. The deployment URL is already pre-configured in the app - **no additional setup needed!**
3.# Step 4: Set Your Google Drive Folder
The script configuration includes the folder structure:
```javascript
const FOLDER_ID = '1EWsCGFT1iqMwUad48t8V_n2Irko-sxGZ';
const FOLDER_STRUCTURE = {
  customerNew: ['CustomerData', 'Beals', 'New'],
  customerProcessedOnly: ['CustomerData', 'Beals', 'Processed', 'CustomerDataOnly'],
  customerAndVendorData: ['CustomerData', 'Beals', 'Processed', 'CustomerAndVendorData'],
  // ... etc
};
```

**To use your own folder:**
1. Open your Google Drive folder
2. Copy the folder ID from the URL: `https://drive.google.com/drive/folders/FOLDER_ID_HERE`
3. Replace the `FOLDER_ID` value in the script
4. Verify your Drive has the required folder structure (see [DEPLOYMENT_INFO.md](DEPLOYMENT_INFO.md))s your folder ID
4. Paste it into the script where it says `const FOLDER_ID = '...';`
5. Save and redeploy

## Troubleshooting

### "Folder not found" Error
- The script is looking for `CustomerData/Beals/` folders
- Make sure this folder path exists in your Drive
- Check the folder names match exactly (case-sensitive)
- See [DEPLOYMENT_INFO.md](DEPLOYMENT_INFO.md) for the required folder structure

### "Authorization required" Error
- When you first run the script, Google asks for permissions
- Click "Review permissions"
- Click your account
- Click "Allow" (the script needs access to your Drive)
- Try again

### Deployment URL doesn't work
- Make sure you selected "Web app" as the deployment type
- Make sure "Anyone" has access
- Try deploying again if it's more than a few minutes old

## Updating the Script

If you need to change the API key or folder ID:
1. Edit the script on script.google.com
2. Change the `AUTHORIZED_API_KEY` or `FOLDER_ID` constant
3. Save the project
4. Click "Deploy" > Click the deployment > Click "Update"
5. You don't need to re-copy the URL

## What Happens Behind the Scenes

1. **Deployment**: You deploy the script on Google's servers
2. **Authorization**: Google grants the script access to your Drive
3. **URL Configured**: The app already knows the deployment URL
4. **Setup Modal**: User opens app and enters API key via setup dialog
5. **Secure Access**: All requests require API key authentication

## Security Notes

- ✅ Only you can see and edit the script on script.google.com
- ✅ The script only accesses the folder you specified
- ✅ Users must have the correct API key to access the app
- ✅ All API requests are authenticated and validated
- ⚠️ Make sure you trust the folder sharing (since it accesses it on your behalf)

## Next Steps

Once deployment is complete:
1. Open the PO Data Processor app
2. The setup modal will appear (first visit only)
3. Enter your API key (from Step 3 of the Google Apps Script setup)
4. Click "Save & Continue"
5. You're all set!

For detailed setup instructions, see [SECURITY_SETUP.md](SECURITY_SETUP.md)

---

**Need help?** Check the error messages in the browser console (F12 > Console tab)
