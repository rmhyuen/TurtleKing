/**
 * Setup and Configuration Management
 * Handles API key setup via localStorage (no config file needed)
 */

const SETUP = {
  STORAGE_KEY: 'po_processor_api_key',
  SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbwknF5fBpZHwy-U3nIhlOA8nWyKKLRx48VfT87XaAithJ3BcpcVx3nIWcY4fXw21dxh/exec',
  
  /**
   * Get API key from localStorage or null if not set
   */
  getApiKey() {
    return localStorage.getItem(this.STORAGE_KEY);
  },
  
  /**
   * Save API key to localStorage
   */
  setApiKey(key) {
    if (!key || key.trim().length === 0) {
      console.error('API key cannot be empty');
      return false;
    }
    localStorage.setItem(this.STORAGE_KEY, key.trim());
    return true;
  },
  
  /**
   * Clear API key from localStorage
   */
  clearApiKey() {
    localStorage.removeItem(this.STORAGE_KEY);
  },
  
  /**
   * Check if API key is configured
   */
  isConfigured() {
    return !!this.getApiKey();
  },
  
  /**
   * Show setup modal if API key not configured
   */
  showSetupIfNeeded() {
    if (!this.isConfigured()) {
      this.showSetupModal();
    }
  },
  
  /**
   * Show setup modal dialog
   */
  showSetupModal() {
    const modal = document.createElement('div');
    modal.id = 'setup-modal';
    modal.className = 'setup-modal';
    modal.innerHTML = `
      <div class="setup-modal-overlay"></div>
      <div class="setup-modal-dialog">
        <div class="setup-modal-header">
          <h2>🔐 Initial Setup Required</h2>
          <p>Enter your API key to enable Google Drive integration</p>
        </div>
        
        <div class="setup-modal-body">
          <div class="setup-form-group">
            <label for="setup-api-key">API Key:</label>
            <input 
              type="password" 
              id="setup-api-key" 
              class="setup-input" 
              placeholder="Enter your API key (925... or your custom key)"
              autocomplete="off"
            />
            <small class="setup-hint">Your API key is stored securely in your browser's local storage. Never shared or sent anywhere except to the Google Apps Script backend.</small>
          </div>
          
          <div class="setup-modal-help">
            <p><strong>Don't have an API key?</strong></p>
            <p>Check the <a href="https://github.com/rmhyuen/TurtleKing/blob/main/Applications/po-data-processor/SECURITY_SETUP.md" target="_blank">SECURITY_SETUP.md</a> file for instructions on how to get your API key.</p>
          </div>
        </div>
        
        <div class="setup-modal-footer">
          <button id="setup-save-btn" class="setup-btn setup-btn-primary">Save & Continue</button>
        </div>
      </div>
    `;
    
    document.body.insertBefore(modal, document.body.firstChild);
    
    // Add styles if not already present
    this.injectStyles();
    
    // Focus on input
    const input = document.getElementById('setup-api-key');
    input.focus();
    
    // Handle enter key
    input.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        this.handleSetupSave();
      }
    });
    
    // Handle save button
    document.getElementById('setup-save-btn').addEventListener('click', () => {
      this.handleSetupSave();
    });
  },
  
  /**
   * Handle setup save
   */
  handleSetupSave() {
    const input = document.getElementById('setup-api-key');
    const apiKey = input.value.trim();
    
    if (!apiKey) {
      alert('Please enter your API key');
      input.focus();
      return;
    }
    
    // Validate API key format (should be UUID-like or hex string)
    if (apiKey.length < 20) {
      alert('API key seems too short. Please verify and try again.');
      input.focus();
      return;
    }
    
    // Save the key
    if (this.setApiKey(apiKey)) {
      // Remove modal and update CONFIG
      const modal = document.getElementById('setup-modal');
      if (modal) {
        modal.remove();
      }
      
      // Update global CONFIG object
      window.CONFIG = {
        scriptUrl: this.SCRIPT_URL,
        apiKey: apiKey,
        maxFileSize: 50 * 1024 * 1024,
        debug: false
      };
      
      console.log('✅ API key configured. Setup complete!');
    } else {
      alert('Failed to save API key. Please try again.');
    }
  },
  
  /**
   * Show settings/configuration panel
   */
  showSettings() {
    const currentKey = this.getApiKey();
    const masked = currentKey ? currentKey.substring(0, 8) + '...' + currentKey.substring(currentKey.length - 4) : 'not set';
    
    const settingsPanel = document.createElement('div');
    settingsPanel.id = 'settings-modal';
    settingsPanel.className = 'setup-modal';
    settingsPanel.innerHTML = `
      <div class="setup-modal-overlay"></div>
      <div class="setup-modal-dialog">
        <div class="setup-modal-header">
          <h2>⚙️ Settings</h2>
        </div>
        
        <div class="setup-modal-body">
          <div class="setup-settings-item">
            <label>API Key Status:</label>
            <p><code>${masked}</code></p>
          </div>
          
          <div class="setup-settings-item">
            <label>Google Apps Script URL:</label>
            <p><code>${this.SCRIPT_URL}</code></p>
          </div>
        </div>
        
        <div class="setup-modal-footer">
          <button id="settings-reconfigure-btn" class="setup-btn setup-btn-secondary">Reconfigure API Key</button>
          <button id="settings-close-btn" class="setup-btn setup-btn-primary">Close</button>
        </div>
      </div>
    `;
    
    document.body.insertBefore(settingsPanel, document.body.firstChild);
    
    document.getElementById('settings-reconfigure-btn').addEventListener('click', () => {
      settingsPanel.remove();
      this.showSetupModal();
    });
    
    document.getElementById('settings-close-btn').addEventListener('click', () => {
      settingsPanel.remove();
    });
  },
  
  /**
   * Inject CSS styles for setup modal
   */
  injectStyles() {
    if (document.getElementById('setup-styles')) return;
    
    const style = document.createElement('style');
    style.id = 'setup-styles';
    style.textContent = `
      .setup-modal {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        z-index: 10000;
        font-family: inherit;
      }
      
      .setup-modal-overlay {
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0, 0, 0, 0.5);
        z-index: -1;
      }
      
      .setup-modal-dialog {
        position: absolute;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        background: white;
        border-radius: 8px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
        width: 90%;
        max-width: 500px;
        max-height: 90vh;
        overflow-y: auto;
        z-index: 10001;
      }
      
      .setup-modal-header {
        padding: 24px 24px 16px;
        border-bottom: 1px solid #e0e0e0;
      }
      
      .setup-modal-header h2 {
        margin: 0 0 8px 0;
        font-size: 20px;
        color: #333;
      }
      
      .setup-modal-header p {
        margin: 0;
        font-size: 14px;
        color: #666;
      }
      
      .setup-modal-body {
        padding: 24px;
      }
      
      .setup-form-group {
        display: flex;
        flex-direction: column;
        margin-bottom: 20px;
      }
      
      .setup-form-group label {
        margin-bottom: 8px;
        font-weight: 600;
        color: #333;
        font-size: 14px;
      }
      
      .setup-input {
        padding: 10px 12px;
        border: 1px solid #ddd;
        border-radius: 4px;
        font-size: 14px;
        font-family: monospace;
        transition: border-color 0.2s;
      }
      
      .setup-input:focus {
        outline: none;
        border-color: #4285f4;
        box-shadow: 0 0 0 3px rgba(66, 133, 244, 0.1);
      }
      
      .setup-hint {
        display: block;
        margin-top: 8px;
        font-size: 12px;
        color: #666;
        line-height: 1.4;
      }
      
      .setup-modal-help {
        background: #f5f5f5;
        padding: 16px;
        border-radius: 4px;
        margin-top: 16px;
      }
      
      .setup-modal-help p {
        margin: 0 0 8px 0;
        font-size: 13px;
        color: #666;
        line-height: 1.4;
      }
      
      .setup-modal-help p:last-child {
        margin-bottom: 0;
      }
      
      .setup-modal-help a {
        color: #4285f4;
        text-decoration: none;
      }
      
      .setup-modal-help a:hover {
        text-decoration: underline;
      }
      
      .setup-modal-footer {
        padding: 16px 24px;
        border-top: 1px solid #e0e0e0;
        display: flex;
        gap: 12px;
        justify-content: flex-end;
      }
      
      .setup-btn {
        padding: 10px 20px;
        border: none;
        border-radius: 4px;
        font-size: 14px;
        font-weight: 600;
        cursor: pointer;
        transition: all 0.2s;
      }
      
      .setup-btn-primary {
        background: #4285f4;
        color: white;
      }
      
      .setup-btn-primary:hover {
        background: #3367d6;
        box-shadow: 0 2px 8px rgba(66, 133, 244, 0.3);
      }
      
      .setup-btn-secondary {
        background: #f5f5f5;
        color: #333;
        border: 1px solid #ddd;
      }
      
      .setup-btn-secondary:hover {
        background: #e8e8e8;
      }
      
      .setup-settings-item {
        margin-bottom: 20px;
      }
      
      .setup-settings-item:last-child {
        margin-bottom: 0;
      }
      
      .setup-settings-item label {
        display: block;
        font-weight: 600;
        color: #333;
        margin-bottom: 8px;
        font-size: 14px;
      }
      
      .setup-settings-item p {
        margin: 0;
        font-size: 13px;
        color: #666;
        word-break: break-all;
      }
      
      .setup-settings-item code {
        display: inline-block;
        background: #f5f5f5;
        padding: 4px 8px;
        border-radius: 3px;
        font-family: monospace;
        font-size: 12px;
      }
    `;
    document.head.appendChild(style);
  }
};

// Initialize setup check when page loads
document.addEventListener('DOMContentLoaded', () => {
  // Initialize CONFIG from localStorage if available
  const apiKey = SETUP.getApiKey();
  window.CONFIG = {
    scriptUrl: SETUP.SCRIPT_URL,
    apiKey: apiKey,
    maxFileSize: 50 * 1024 * 1024,
    debug: false
  };
  
  // Show setup if API key not configured
  if (!apiKey) {
    SETUP.showSetupIfNeeded();
  } else {
    console.log('✅ API key loaded from browser storage');
  }
});
