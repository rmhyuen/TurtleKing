# Security Setup Guide - Updated

## 🎉 Simplified Setup with localStorage

The app now uses **browser localStorage** for API key management. **No config files to download or manage!**

---

## Quick Start (2 minutes)

1. **Open the app** at [https://rmhyuen.github.io/TurtleKing/apps/po-data-processor](https://rmhyuen.github.io/TurtleKing/apps/po-data-processor)

2. **A setup dialog appears** automatically asking for your API key

3. **Enter your API key:**

4. **Click "Save & Continue"**

5. **Done!** Your API key is securely stored in your browser (never sent anywhere except to the Google Apps Script backend)

---

## How It Works

- **First visit:** Setup modal appears → you enter API key once
- **Subsequent visits:** API key automatically loaded from browser storage
- **Settings button (⚙️):** Top right corner - view or reconfigure your API key
- **Stored locally:** Only in your browser, never on servers or github

---

## Changing Your API Key

1. Click the settings button (⚙️) at the top right
2. Click "Reconfigure API Key"
3. Enter your new API key
4. Click "Save & Continue"

---

## If You've Changed Your API Key in Google Apps Script

1. Open your Google Apps Script: https://script.google.com
2. Update the `AUTHORIZED_API_KEY` constant:
   ```javascript
   const AUTHORIZED_API_KEY = 'your-new-key-here';
   ```
3. **Deploy** → **Manage Deployments** → Edit → **Redeploy**
4. Return to the PO Data Processor
5. Click settings (⚙️) → "Reconfigure API Key" → enter your new key

---

## ✅ Security Benefits

✅ **No config files** - nothing to accidentally commit to git  
✅ **Browser-only storage** - API key never leaves your machine except for legitimate API calls  
✅ **Secure transmission** - HTTPS only for all communication  
✅ **User-controlled** - you can reset/change the key anytime  
✅ **Works on GitHub Pages** - no server configuration needed  

---

## ⚠️ Important Notes

- **Clear browser storage?** If you clear your browser's site data, you'll need to re-enter your API key
- **Multiple browsers/devices?** You'll need to set the API key on each browser (it's stored locally, not synced)
- **Share a device?** Other users won't see your API key (it's browser-private)

---

## 🆘 Troubleshooting

**"Unauthorized" error on file uploads/downloads**
- Double-check that your API key matches the one in your Google Apps Script
- Check browser DevTools → Console tab for error messages
- Try clicking Settings (⚙️) and reconfiguring the API key

**Setup modal keeps appearing**
- Your browser might have private/incognito mode enabled (localStorage disabled)
- Try using a regular browsing window instead
- Check browser settings to allow localStorage for this site

**I lost my API key**
- Check your Google Apps Script project for the `AUTHORIZED_API_KEY` constant
- Or generate a new one: [SECURITY.md](SECURITY.md) has instructions

---

That's it! Your app is now configured and ready to use.
