// Content script to interact with web pages
console.log('Oracle Extension content script loaded');

// Create toast notification styles
const createToastStyles = () => {
  if (document.getElementById('oracle-toast-styles')) return;

  const styles = document.createElement('style');
  styles.id = 'oracle-toast-styles';
  styles.textContent = `
    .oracle-toast {
      position: fixed !important;
      top: 20px !important;
      right: 20px !important;
      background: #2c3e50 !important;
      color: white !important;
      padding: 15px 20px !important;
      border-radius: 8px !important;
      box-shadow: 0 4px 12px rgba(0,0,0,0.3) !important;
      z-index: 999999 !important;
      font-family: -apple-system, BlinkMacSystemFont, sans-serif !important;
      font-size: 14px !important;
      max-width: 350px !important;
      animation: oracleSlideIn 0.3s ease-out !important;
      border-left: 4px solid #3498db !important;
    }
    
    .oracle-toast.success {
      border-left-color: #27ae60 !important;
    }
    
    .oracle-toast.error {
      border-left-color: #e74c3c !important;
    }
    
    .oracle-toast-title {
      font-weight: bold !important;
      margin-bottom: 5px !important;
      color: #3498db !important;
    }
    
    .oracle-toast.success .oracle-toast-title {
      color: #27ae60 !important;
    }
    
    .oracle-toast.error .oracle-toast-title {
      color: #e74c3c !important;
    }
    
    .oracle-toast-close {
      position: absolute !important;
      top: 8px !important;
      right: 10px !important;
      background: none !important;
      border: none !important;
      color: #bdc3c7 !important;
      cursor: pointer !important;
      font-size: 16px !important;
      padding: 0 !important;
      width: 20px !important;
      height: 20px !important;
    }
    
    .oracle-toast-close:hover {
      color: white !important;
    }
    
    @keyframes oracleSlideIn {
      from {
        transform: translateX(100%) !important;
        opacity: 0 !important;
      }
      to {
        transform: translateX(0) !important;
        opacity: 1 !important;
      }
    }
    
    @keyframes oracleSlideOut {
      from {
        transform: translateX(0) !important;
        opacity: 1 !important;
      }
      to {
        transform: translateX(100%) !important;
        opacity: 0 !important;
      }
    }
    
    .oracle-alert-overlay {
      position: fixed !important;
      top: 0 !important;
      left: 0 !important;
      right: 0 !important;
      bottom: 0 !important;
      background: rgba(0, 0, 0, 0.6) !important;
      z-index: 9999999 !important;
      display: flex !important;
      align-items: center !important;
      justify-content: center !important;
      animation: fadeIn 0.2s ease-out !important;
    }
    
    .oracle-alert-box {
      background: white !important;
      border-radius: 12px !important;
      box-shadow: 0 8px 32px rgba(0,0,0,0.3) !important;
      max-width: 600px !important;
      width: 90% !important;
      max-height: 80vh !important;
      overflow: hidden !important;
      display: flex !important;
      flex-direction: column !important;
      animation: scaleIn 0.3s ease-out !important;
    }
    
    .oracle-alert-header {
      background: linear-gradient(45deg, #667eea, #764ba2) !important;
      color: white !important;
      padding: 16px 20px !important;
      font-size: 18px !important;
      font-weight: 600 !important;
      display: flex !important;
      align-items: center !important;
      justify-content: space-between !important;
    }
    
    .oracle-alert-content {
      padding: 20px !important;
      overflow-y: auto !important;
      flex: 1 !important;
      font-family: -apple-system, BlinkMacSystemFont, sans-serif !important;
      font-size: 14px !important;
      line-height: 1.6 !important;
      color: #2c3e50 !important;
      white-space: pre-wrap !important;
    }
    
    .oracle-alert-footer {
      padding: 16px 20px !important;
      border-top: 1px solid #e1e8ed !important;
      display: flex !important;
      justify-content: flex-end !important;
      gap: 10px !important;
    }
    
    .oracle-alert-btn {
      padding: 10px 20px !important;
      border-radius: 8px !important;
      font-size: 14px !important;
      font-weight: 600 !important;
      cursor: pointer !important;
      transition: all 0.2s !important;
      border: none !important;
    }
    
    .oracle-alert-btn-close {
      background: linear-gradient(45deg, #667eea, #764ba2) !important;
      color: white !important;
    }
    
    .oracle-alert-btn-close:hover {
      transform: translateY(-2px) !important;
      box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4) !important;
    }
    
    .oracle-alert-btn-copy {
      background: #ecf0f1 !important;
      color: #2c3e50 !important;
    }
    
    .oracle-alert-btn-copy:hover {
      background: #bdc3c7 !important;
    }
    
    .oracle-alert-btn-copy.copied {
      background: #27ae60 !important;
      color: white !important;
    }
    
    @keyframes fadeIn {
      from { opacity: 0; }
      to { opacity: 1; }
    }
    
    @keyframes scaleIn {
      from {
        transform: scale(0.9);
        opacity: 0;
      }
      to {
        transform: scale(1);
        opacity: 1;
      }
    }
    
    /* Priority Task Modal Styles */
    .oracle-priority-modal-overlay {
      position: fixed !important;
      top: 0 !important;
      left: 0 !important;
      right: 0 !important;
      bottom: 0 !important;
      background: rgba(0, 0, 0, 0.6) !important;
      z-index: 9999999 !important;
      display: flex !important;
      align-items: center !important;
      justify-content: center !important;
      animation: fadeIn 0.2s ease-out !important;
    }
    
    .oracle-priority-modal {
      background: white !important;
      border-radius: 12px !important;
      box-shadow: 0 8px 32px rgba(0,0,0,0.3) !important;
      max-width: 500px !important;
      width: 90% !important;
      max-height: 80vh !important;
      overflow: hidden !important;
      display: flex !important;
      flex-direction: column !important;
      animation: scaleIn 0.3s ease-out !important;
    }
    
    .oracle-priority-modal-header {
      background: linear-gradient(45deg, #667eea, #764ba2) !important;
      color: white !important;
      padding: 16px 20px !important;
      font-size: 18px !important;
      font-weight: 600 !important;
      display: flex !important;
      align-items: center !important;
      justify-content: space-between !important;
    }
    
    .oracle-priority-modal-content {
      padding: 20px !important;
      overflow-y: auto !important;
      flex: 1 !important;
      font-family: -apple-system, BlinkMacSystemFont, sans-serif !important;
    }
    
    .oracle-priority-modal-field {
      margin-bottom: 16px !important;
    }
    
    .oracle-priority-modal-label {
      display: block !important;
      font-size: 13px !important;
      font-weight: 600 !important;
      color: #2c3e50 !important;
      margin-bottom: 6px !important;
    }
    
    .oracle-priority-modal-input {
      width: 100% !important;
      padding: 10px 12px !important;
      border: 1px solid #cbd5e0 !important;
      border-radius: 6px !important;
      font-size: 14px !important;
      font-family: -apple-system, BlinkMacSystemFont, sans-serif !important;
      transition: border-color 0.2s !important;
      box-sizing: border-box !important;
    }
    
    .oracle-priority-modal-input:focus {
      outline: none !important;
      border-color: #667eea !important;
      box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1) !important;
    }
    
    .oracle-priority-modal-input:disabled {
      background: #f7fafc !important;
      color: #4a5568 !important;
      cursor: not-allowed !important;
    }
    
    .oracle-priority-modal-tags-container {
      border: 1px solid #cbd5e0 !important;
      border-radius: 6px !important;
      padding: 8px !important;
      min-height: 42px !important;
      display: flex !important;
      flex-wrap: wrap !important;
      gap: 6px !important;
      align-items: center !important;
      background: white !important;
      cursor: text !important;
      box-sizing: border-box !important;
      width: 100% !important;
    }
    
    .oracle-priority-modal-tags-container:focus-within {
      border-color: #667eea !important;
      box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1) !important;
    }
    
    .oracle-priority-modal-tag {
      display: inline-flex !important;
      align-items: center !important;
      background: linear-gradient(45deg, #667eea, #764ba2) !important;
      color: white !important;
      padding: 4px 10px !important;
      border-radius: 4px !important;
      font-size: 13px !important;
      gap: 6px !important;
    }
    
    .oracle-priority-modal-tag-remove {
      background: rgba(255, 255, 255, 0.3) !important;
      border: none !important;
      color: white !important;
      cursor: pointer !important;
      border-radius: 50% !important;
      width: 16px !important;
      height: 16px !important;
      display: flex !important;
      align-items: center !important;
      justify-content: center !important;
      font-size: 12px !important;
      padding: 0 !important;
      transition: background 0.2s !important;
    }
    
    .oracle-priority-modal-tag-remove:hover {
      background: rgba(255, 255, 255, 0.5) !important;
    }
    
    .oracle-priority-modal-tag-input {
      flex: 1 !important;
      min-width: 120px !important;
      border: none !important;
      outline: none !important;
      font-size: 14px !important;
      font-family: -apple-system, BlinkMacSystemFont, sans-serif !important;
      padding: 4px !important;
    }
    
    .oracle-priority-modal-footer {
      padding: 16px 20px !important;
      border-top: 1px solid #e1e8ed !important;
      display: flex !important;
      justify-content: flex-end !important;
      gap: 10px !important;
    }
    
    .oracle-priority-modal-btn {
      padding: 10px 20px !important;
      border-radius: 8px !important;
      font-size: 14px !important;
      font-weight: 600 !important;
      cursor: pointer !important;
      transition: all 0.2s !important;
      border: none !important;
    }
    
    .oracle-priority-modal-btn-cancel {
      background: #ecf0f1 !important;
      color: #2c3e50 !important;
    }
    
    .oracle-priority-modal-btn-cancel:hover {
      background: #bdc3c7 !important;
    }
    
    .oracle-priority-modal-btn-submit {
      background: linear-gradient(45deg, #667eea, #764ba2) !important;
      color: white !important;
    }
    
    .oracle-priority-modal-btn-submit:hover {
      transform: translateY(-2px) !important;
      box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4) !important;
    }
    
    .oracle-priority-modal-btn-submit:disabled {
      opacity: 0.5 !important;
      cursor: not-allowed !important;
      transform: none !important;
    }
  `;
  document.head.appendChild(styles);
};

// Show toast notification
const showToast = (message, type = 'success', title = 'Oracle', duration = 8000) => {
  createToastStyles();

  // Remove existing toast if any
  const existingToast = document.getElementById('oracle-toast');
  if (existingToast) {
    existingToast.remove();
  }

  // Create toast element
  const toast = document.createElement('div');
  toast.id = 'oracle-toast';
  toast.className = `oracle-toast ${type}`;
  toast.innerHTML = `
    <button class="oracle-toast-close">×</button>
    <div class="oracle-toast-title">${title}</div>
    <div>${message}</div>
  `;

  // Add close functionality
  toast.querySelector('.oracle-toast-close').addEventListener('click', () => {
    toast.style.animation = 'oracleSlideOut 0.3s ease-out';
    setTimeout(() => toast.remove(), 300);
  });

  // Add to page
  document.body.appendChild(toast);

  // Auto-remove after duration
  setTimeout(() => {
    if (toast.parentNode) {
      toast.style.animation = 'oracleSlideOut 0.3s ease-out';
      setTimeout(() => {
        if (toast.parentNode) toast.remove();
      }, 300);
    }
  }, duration);
};

// Show alert dialog
const showAlert = (message, title = 'Oracle') => {
  createToastStyles();

  // Remove existing alert if any
  const existingAlert = document.getElementById('oracle-alert-overlay');
  if (existingAlert) {
    existingAlert.remove();
  }

  // Create alert overlay
  const overlay = document.createElement('div');
  overlay.id = 'oracle-alert-overlay';
  overlay.className = 'oracle-alert-overlay';

  // Create alert box
  const alertBox = document.createElement('div');
  alertBox.className = 'oracle-alert-box';
  alertBox.innerHTML = `
    <div class="oracle-alert-header">
      ${title}
    </div>
    <div class="oracle-alert-content">${message}</div>
    <div class="oracle-alert-footer">
      <button class="oracle-alert-btn oracle-alert-btn-copy">Copy</button>
      <button class="oracle-alert-btn oracle-alert-btn-close">Close</button>
    </div>
  `;

  overlay.appendChild(alertBox);
  document.body.appendChild(overlay);

  // Add close functionality
  const closeBtn = alertBox.querySelector('.oracle-alert-btn-close');
  closeBtn.addEventListener('click', () => {
    overlay.style.animation = 'fadeOut 0.2s ease-out';
    setTimeout(() => overlay.remove(), 200);
  });

  // Close on overlay click
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) {
      overlay.style.animation = 'fadeOut 0.2s ease-out';
      setTimeout(() => overlay.remove(), 200);
    }
  });

  // Add copy functionality
  const copyBtn = alertBox.querySelector('.oracle-alert-btn-copy');
  copyBtn.addEventListener('click', () => {
    navigator.clipboard.writeText(message);
    copyBtn.textContent = 'Copied!';
    copyBtn.classList.add('copied');
    setTimeout(() => {
      copyBtn.textContent = 'Copy';
      copyBtn.classList.remove('copied');
    }, 2000);
  });
};

// Show Priority Task modal
const showPriorityTaskModal = (data, selectedText, sourceUrl) => {
  console.log('showPriorityTaskModal called!');
  console.log('Data:', data);
  console.log('Selected text:', selectedText);
  console.log('Source URL:', sourceUrl);

  createToastStyles();

  // Remove existing modal if any
  const existingModal = document.getElementById('oracle-priority-modal-overlay');
  if (existingModal) {
    existingModal.remove();
  }

  // Parse response data
  let modalData = {};
  if (Array.isArray(data) && data.length > 0) {
    modalData = data[0];
  } else {
    modalData = data;
  }

  const title = modalData.title || modalData.task_title || '';
  const suggestedTags = modalData.tags || modalData.suggested_tags || [];
  const WEBHOOK_URL = 'https://n8n-kqq5.onrender.com/webhook/d60909cd-7ae4-431e-9c4f-0c9cacfa20ea';

  // Create modal overlay
  const overlay = document.createElement('div');
  overlay.id = 'oracle-priority-modal-overlay';
  overlay.className = 'oracle-priority-modal-overlay';

  // Create modal box
  const modalBox = document.createElement('div');
  modalBox.className = 'oracle-priority-modal';

  // Create tags array to manage
  let tags = [...suggestedTags];

  // Build tags HTML
  const buildTagsHTML = () => {
    return tags.map(tag => `
      <span class="oracle-priority-modal-tag">
        ${escapeHtml(tag)}
        <button class="oracle-priority-modal-tag-remove" data-tag="${escapeHtml(tag)}">×</button>
      </span>
    `).join('');
  };

  // Helper function to escape HTML
  function escapeHtml(unsafe) {
    return (unsafe || '')
      .toString()
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  modalBox.innerHTML = `
    <div class="oracle-priority-modal-header">
      Create Priority Task
    </div>
    <div class="oracle-priority-modal-content">
      <div class="oracle-priority-modal-field">
        <label class="oracle-priority-modal-label">Title</label>
        <input 
          type="text" 
          class="oracle-priority-modal-input" 
          id="oracle-priority-modal-title"
          value="${escapeHtml(title)}"
          placeholder="Enter task title..."
        />
      </div>
      
      <div class="oracle-priority-modal-field">
        <label class="oracle-priority-modal-label">Message</label>
        <input 
          type="text" 
          class="oracle-priority-modal-input" 
          id="oracle-priority-modal-message"
          value="${escapeHtml(selectedText)}"
          disabled
        />
      </div>
      
      <div class="oracle-priority-modal-field">
        <label class="oracle-priority-modal-label">Tags</label>
        <div class="oracle-priority-modal-tags-container" id="oracle-priority-modal-tags">
          ${buildTagsHTML()}
          <input 
            type="text" 
            class="oracle-priority-modal-tag-input" 
            id="oracle-priority-modal-tag-input"
            placeholder="Add tag..."
          />
        </div>
      </div>
    </div>
    <div class="oracle-priority-modal-footer">
      <button class="oracle-priority-modal-btn oracle-priority-modal-btn-cancel" id="oracle-priority-modal-cancel">Cancel</button>
      <button class="oracle-priority-modal-btn oracle-priority-modal-btn-submit" id="oracle-priority-modal-submit">Submit</button>
    </div>
  `;

  overlay.appendChild(modalBox);
  document.body.appendChild(overlay);

  // Get elements
  const titleInput = document.getElementById('oracle-priority-modal-title');
  const tagsContainer = document.getElementById('oracle-priority-modal-tags');
  const tagInput = document.getElementById('oracle-priority-modal-tag-input');
  const cancelBtn = document.getElementById('oracle-priority-modal-cancel');
  const submitBtn = document.getElementById('oracle-priority-modal-submit');

  // Update tags display
  const updateTagsDisplay = () => {
    const tagsHTML = buildTagsHTML();
    tagsContainer.innerHTML = tagsHTML + `
      <input 
        type="text" 
        class="oracle-priority-modal-tag-input" 
        id="oracle-priority-modal-tag-input"
        placeholder="Add tag..."
      />
    `;

    // Re-attach event listeners to remove buttons
    tagsContainer.querySelectorAll('.oracle-priority-modal-tag-remove').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const tagToRemove = btn.getAttribute('data-tag');
        tags = tags.filter(t => t !== tagToRemove);
        updateTagsDisplay();
      });
    });

    // Re-focus on input
    const newTagInput = document.getElementById('oracle-priority-modal-tag-input');
    newTagInput.focus();

    // Re-attach input event listener
    newTagInput.addEventListener('keydown', handleTagInput);
  };

  // Handle tag input
  const handleTagInput = (e) => {
    if (e.key === 'Enter' || e.key === ',') {
      e.preventDefault();
      const value = e.target.value.trim();
      if (value && !tags.includes(value)) {
        tags.push(value);
        updateTagsDisplay();
      }
      e.target.value = '';
    }
  };

  // Initial tag input listener
  tagInput.addEventListener('keydown', handleTagInput);

  // Click on tags container to focus input
  tagsContainer.addEventListener('click', () => {
    const currentTagInput = document.getElementById('oracle-priority-modal-tag-input');
    if (currentTagInput) {
      currentTagInput.focus();
    }
  });

  // Initial tag remove listeners
  tagsContainer.querySelectorAll('.oracle-priority-modal-tag-remove').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const tagToRemove = btn.getAttribute('data-tag');
      tags = tags.filter(t => t !== tagToRemove);
      updateTagsDisplay();
    });
  });

  // Close modal function
  const closeModal = () => {
    overlay.style.animation = 'fadeOut 0.2s ease-out';
    setTimeout(() => overlay.remove(), 200);
  };

  // Cancel button
  cancelBtn.addEventListener('click', closeModal);

  // Close on overlay click
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) {
      closeModal();
    }
  });

  // Submit button
  submitBtn.addEventListener('click', async () => {
    const finalTitle = titleInput.value.trim();

    if (!finalTitle) {
      titleInput.focus();
      titleInput.style.borderColor = '#e74c3c';
      setTimeout(() => {
        titleInput.style.borderColor = '#cbd5e0';
      }, 1500);
      return;
    }

    // Disable submit button
    submitBtn.disabled = true;
    submitBtn.textContent = 'Creating...';

    // Get user data from storage
    try {
      const result = await chrome.storage.local.get(['oracle_user_data']);
      const userData = result.oracle_user_data || {};

      // Prepare payload
      const payload = {
        action: 'create_priority_task',
        intent: 'priority_task_creation',
        task_title: finalTitle,
        task_text: selectedText,
        tags: tags,
        sourceUrl: sourceUrl,
        priority: true,
        timestamp: new Date().toISOString(),
        source: 'oracle-chrome-extension-modal',
        user_id: userData.userId,
        authenticated: true
      };

      // Send to background script to make the webhook call (avoids CORS)
      const response = await chrome.runtime.sendMessage({
        action: 'makeWebhookCall',
        payload: payload
      });

      if (response && response.success) {
        // Close modal
        closeModal();

        // Show success toast
        showToast('Priority task created successfully!', 'success', 'Oracle - Priority Task');

        // Request todo list refresh
        chrome.runtime.sendMessage({ action: 'refreshTodoList' }).catch(() => {
          console.log('Could not send refresh message to background');
        });
      } else {
        throw new Error(response?.error || 'Failed to create task');
      }
    } catch (error) {
      console.error('Error creating priority task:', error);
      showToast(`Error: ${error.message}`, 'error', 'Oracle - Error');
      submitBtn.disabled = false;
      submitBtn.textContent = 'Submit';
    }
  });

  // Focus on title input
  titleInput.focus();
};

// Show Bookmark modal
const showBookmarkModal = (data, bookmarkUrl, pageTitle, selectedText) => {
  console.log('showBookmarkModal called!');
  console.log('Data:', data);
  console.log('Bookmark URL:', bookmarkUrl);
  console.log('Page title:', pageTitle);

  createToastStyles();

  // Remove existing modal if any
  const existingModal = document.getElementById('oracle-bookmark-modal-overlay');
  if (existingModal) {
    existingModal.remove();
  }

  // Parse response data
  let modalData = {};
  if (Array.isArray(data) && data.length > 0) {
    modalData = data[0];
  } else {
    modalData = data;
  }

  const title = modalData.title || modalData.suggested_title || pageTitle || '';

  // Use bookmarkUrl from response if present, otherwise use current page URL
  const url = modalData.bookmarkUrl || modalData.url || bookmarkUrl || '';

  let suggestedTags = modalData.tags || modalData.suggested_tags || [];

  // Parse tags if they're a JSON string
  if (typeof suggestedTags === 'string') {
    try {
      suggestedTags = JSON.parse(suggestedTags);
    } catch (e) {
      suggestedTags = [];
    }
  }

  const WEBHOOK_URL = 'https://n8n-kqq5.onrender.com/webhook/d60909cd-7ae4-431e-9c4f-0c9cacfa20ea';

  // Create modal overlay
  const overlay = document.createElement('div');
  overlay.id = 'oracle-bookmark-modal-overlay';
  overlay.className = 'oracle-priority-modal-overlay';

  // Create modal box
  const modalBox = document.createElement('div');
  modalBox.className = 'oracle-priority-modal';

  // Create tags array to manage (start empty)
  let tags = [];

  // Helper function to escape HTML
  function escapeHtml(unsafe) {
    return (unsafe || '')
      .toString()
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  // Build tags HTML
  const buildTagsHTML = () => {
    return tags.map(tag => `
      <span class="oracle-priority-modal-tag">
        ${escapeHtml(tag)}
        <button class="oracle-priority-modal-tag-remove" data-tag="${escapeHtml(tag)}">×</button>
      </span>
    `).join('');
  };

  // Build suggested tags dropdown options
  const buildSuggestedTagsOptions = () => {
    return suggestedTags
      .filter(tag => !tags.includes(tag)) // Don't show already selected tags
      .map(tag => `<option value="${escapeHtml(tag)}">${escapeHtml(tag)}</option>`)
      .join('');
  };

  modalBox.innerHTML = `
    <div class="oracle-priority-modal-header">
      Create Bookmark
    </div>
    <div class="oracle-priority-modal-content">
      <div class="oracle-priority-modal-field">
        <label class="oracle-priority-modal-label">Title</label>
        <input 
          type="text" 
          class="oracle-priority-modal-input" 
          id="oracle-bookmark-modal-title"
          value="${escapeHtml(title)}"
          placeholder="Enter bookmark title..."
        />
      </div>
      
      <div class="oracle-priority-modal-field">
        <label class="oracle-priority-modal-label">URL</label>
        <input 
          type="text" 
          class="oracle-priority-modal-input" 
          id="oracle-bookmark-modal-url"
          value="${escapeHtml(url)}"
          placeholder="Enter URL..."
        />
      </div>
      
      <div class="oracle-priority-modal-field">
        <label class="oracle-priority-modal-label">Tags</label>
        <div class="oracle-priority-modal-tags-container" id="oracle-bookmark-modal-tags">
          ${buildTagsHTML()}
          <div style="position: relative; display: inline-block; flex: 1; min-width: 120px;">
            <input 
              type="text" 
              class="oracle-priority-modal-tag-input" 
              id="oracle-bookmark-modal-tag-input"
              placeholder="Add tag..."
              list="oracle-bookmark-suggested-tags"
              autocomplete="off"
            />
            <datalist id="oracle-bookmark-suggested-tags">
              ${buildSuggestedTagsOptions()}
            </datalist>
          </div>
        </div>
        ${suggestedTags.length > 0 ? `
          <div style="margin-top: 8px; font-size: 12px; color: #7f8c8d;">
            <span style="font-weight: 500;">Suggested:</span>
            ${suggestedTags.map(tag => `
              <span class="oracle-bookmark-suggested-tag" data-tag="${escapeHtml(tag)}" style="display: inline-block; background: rgba(102, 126, 234, 0.1); color: #667eea; padding: 2px 6px; border-radius: 4px; font-size: 11px; margin-left: 4px; cursor: pointer; border: 1px solid rgba(102, 126, 234, 0.3); transition: all 0.2s;">
                ${escapeHtml(tag)}
              </span>
            `).join('')}
          </div>
        ` : ''}
      </div>
    </div>
    <div class="oracle-priority-modal-footer">
      <button class="oracle-priority-modal-btn oracle-priority-modal-btn-cancel" id="oracle-bookmark-modal-cancel">Cancel</button>
      <button class="oracle-priority-modal-btn oracle-priority-modal-btn-submit" id="oracle-bookmark-modal-submit">Submit</button>
    </div>
  `;

  overlay.appendChild(modalBox);
  document.body.appendChild(overlay);

  // Get elements
  const titleInput = document.getElementById('oracle-bookmark-modal-title');
  const urlInput = document.getElementById('oracle-bookmark-modal-url');
  const tagsContainer = document.getElementById('oracle-bookmark-modal-tags');
  const tagInput = document.getElementById('oracle-bookmark-modal-tag-input');
  const cancelBtn = document.getElementById('oracle-bookmark-modal-cancel');
  const submitBtn = document.getElementById('oracle-bookmark-modal-submit');

  // Update tags display
  const updateTagsDisplay = () => {
    const tagsHTML = buildTagsHTML();
    tagsContainer.innerHTML = tagsHTML + `
      <div style="position: relative; display: inline-block; flex: 1; min-width: 120px;">
        <input 
          type="text" 
          class="oracle-priority-modal-tag-input" 
          id="oracle-bookmark-modal-tag-input"
          placeholder="Add tag..."
          list="oracle-bookmark-suggested-tags"
          autocomplete="off"
        />
        <datalist id="oracle-bookmark-suggested-tags">
          ${buildSuggestedTagsOptions()}
        </datalist>
      </div>
    `;

    // Re-attach event listeners to remove buttons
    tagsContainer.querySelectorAll('.oracle-priority-modal-tag-remove').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const tagToRemove = btn.getAttribute('data-tag');
        tags = tags.filter(t => t !== tagToRemove);
        updateTagsDisplay();
      });
    });

    // Re-attach suggested tag click listeners
    document.querySelectorAll('.oracle-bookmark-suggested-tag').forEach(span => {
      span.addEventListener('click', (e) => {
        e.stopPropagation();
        const tagToAdd = span.getAttribute('data-tag');
        if (!tags.includes(tagToAdd)) {
          tags.push(tagToAdd);
          updateTagsDisplay();
        }
      });

      // Hover effects
      span.addEventListener('mouseenter', () => {
        span.style.background = 'linear-gradient(45deg, rgba(102, 126, 234, 0.2), rgba(118, 75, 162, 0.2))';
        span.style.transform = 'translateY(-1px)';
      });
      span.addEventListener('mouseleave', () => {
        span.style.background = 'rgba(102, 126, 234, 0.1)';
        span.style.transform = 'translateY(0)';
      });
    });

    // Re-focus on input
    const newTagInput = document.getElementById('oracle-bookmark-modal-tag-input');
    newTagInput.focus();

    // Re-attach input event listener
    newTagInput.addEventListener('keydown', handleTagInput);
    newTagInput.addEventListener('input', handleTagInputChange);
  };

  // Handle tag input
  const handleTagInput = (e) => {
    if (e.key === 'Enter' || e.key === ',') {
      e.preventDefault();
      const value = e.target.value.trim();
      if (value && !tags.includes(value)) {
        tags.push(value);
        updateTagsDisplay();
      }
      e.target.value = '';
    } else if (e.key === 'Backspace' && e.target.value === '' && tags.length > 0) {
      // Remove last tag if backspace on empty input
      tags.pop();
      updateTagsDisplay();
    }
  };

  // Handle tag input change (for datalist selection)
  const handleTagInputChange = (e) => {
    const value = e.target.value.trim();
    // Check if value matches a suggested tag
    if (suggestedTags.includes(value) && !tags.includes(value)) {
      tags.push(value);
      updateTagsDisplay();
    }
  };

  // Initial tag input listener
  tagInput.addEventListener('keydown', handleTagInput);
  tagInput.addEventListener('input', handleTagInputChange);

  // Click on tags container to focus input
  tagsContainer.addEventListener('click', () => {
    const currentTagInput = document.getElementById('oracle-bookmark-modal-tag-input');
    if (currentTagInput) {
      currentTagInput.focus();
    }
  });

  // Suggested tag click listeners
  document.querySelectorAll('.oracle-bookmark-suggested-tag').forEach(span => {
    span.addEventListener('click', (e) => {
      e.stopPropagation();
      const tagToAdd = span.getAttribute('data-tag');
      if (!tags.includes(tagToAdd)) {
        tags.push(tagToAdd);
        updateTagsDisplay();
      }
    });

    // Hover effects
    span.addEventListener('mouseenter', () => {
      span.style.background = 'linear-gradient(45deg, rgba(102, 126, 234, 0.2), rgba(118, 75, 162, 0.2))';
      span.style.transform = 'translateY(-1px)';
    });
    span.addEventListener('mouseleave', () => {
      span.style.background = 'rgba(102, 126, 234, 0.1)';
      span.style.transform = 'translateY(0)';
    });
  });

  // Close modal function
  const closeModal = () => {
    overlay.style.animation = 'fadeOut 0.2s ease-out';
    setTimeout(() => overlay.remove(), 200);
  };

  // Cancel button
  cancelBtn.addEventListener('click', closeModal);

  // Close on overlay click
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) {
      closeModal();
    }
  });

  // Submit button
  submitBtn.addEventListener('click', async () => {
    const finalTitle = titleInput.value.trim();
    const finalUrl = urlInput.value.trim();

    if (!finalTitle || !finalUrl) {
      if (!finalTitle) {
        titleInput.focus();
        titleInput.style.borderColor = '#e74c3c';
      }
      if (!finalUrl) {
        urlInput.focus();
        urlInput.style.borderColor = '#e74c3c';
      }
      setTimeout(() => {
        titleInput.style.borderColor = '#cbd5e0';
        urlInput.style.borderColor = '#cbd5e0';
      }, 1500);
      return;
    }

    // Disable submit button
    submitBtn.disabled = true;
    submitBtn.textContent = 'Creating...';

    // Get user data from storage
    try {
      const result = await chrome.storage.local.get(['oracle_user_data']);
      const userData = result.oracle_user_data || {};

      // Prepare payload - send tags as array for cleaner data
      const payload = {
        action: 'create_bookmark',
        intent: 'bookmark_creation',
        title: finalTitle,
        bookmarkUrl: finalUrl,
        bookmarkText: selectedText || finalTitle,
        tags: tags, // Send as array
        timestamp: new Date().toISOString(),
        source: 'oracle-chrome-extension-modal',
        user_id: userData.userId,
        authenticated: true
      };

      // Send to background script to make the webhook call (avoids CORS)
      const response = await chrome.runtime.sendMessage({
        action: 'makeWebhookCall',
        payload: payload
      });

      if (response && response.success) {
        // Close modal
        closeModal();

        // Show success toast
        showToast('Bookmark created successfully!', 'success', 'Oracle - Bookmark');

        // Request todo list refresh (this will also refresh bookmarks if side panel is open)
        chrome.runtime.sendMessage({ action: 'refreshTodoList' }).catch(() => {
          console.log('Could not send refresh message to background');
        });
      } else {
        throw new Error(response?.error || 'Failed to create bookmark');
      }
    } catch (error) {
      console.error('Error creating bookmark:', error);
      showToast(`Error: ${error.message}`, 'error', 'Oracle - Error');
      submitBtn.disabled = false;
      submitBtn.textContent = 'Submit';
    }
  });

  // Focus on title input
  titleInput.focus();
};

// Listen for messages from popup and background script
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'getPageData') {
    try {
      const pageData = {
        url: window.location.href,
        title: document.title,
        selectedText: window.getSelection().toString().trim(),
        timestamp: new Date().toISOString(),
        domain: window.location.hostname,
        path: window.location.pathname,
        metaDescription: document.querySelector('meta[name="description"]')?.content || '',
        headings: Array.from(document.querySelectorAll('h1, h2, h3')).map(h => h.textContent.trim()).slice(0, 5),
        textLength: document.body.innerText.length,
        hasForm: document.forms.length > 0,
        imageCount: document.images.length
      };

      sendResponse(pageData);
    } catch (error) {
      console.error('Error getting page data:', error);
      sendResponse({ error: error.message });
    }
  } else if (request.action === 'showToast') {
    // Show toast notification from background script
    showToast(request.message, request.type || 'success', request.title || 'Oracle');
    sendResponse({ success: true });
  } else if (request.action === 'showAlert') {
    // Show alert dialog from background script
    showAlert(request.message, request.title || 'Oracle');
    sendResponse({ success: true });
  } else if (request.action === 'showPriorityTaskModal') {
    // Show Priority Task modal
    console.log('Received showPriorityTaskModal message!');
    console.log('Modal data:', request.data);
    console.log('Selected text:', request.selectedText);
    showPriorityTaskModal(request.data, request.selectedText, request.sourceUrl);
    sendResponse({ success: true });
  } else if (request.action === 'showBookmarkModal') {
    // Show Bookmark modal
    console.log('Received showBookmarkModal message!');
    console.log('Modal data:', request.data);
    console.log('Bookmark URL:', request.bookmarkUrl);
    console.log('Page title:', request.pageTitle);
    showBookmarkModal(request.data, request.bookmarkUrl, request.pageTitle, request.selectedText);
    sendResponse({ success: true });
  }

  return true; // Keep message channel open for async response
});

// Optional: Add right-click context menu trigger
document.addEventListener('contextmenu', (event) => {
  // Store the clicked element for potential use
  chrome.storage.local.set({
    lastClickedElement: {
      tagName: event.target.tagName,
      className: event.target.className,
      id: event.target.id,
      text: event.target.textContent?.substring(0, 100) || ''
    }
  });
});
