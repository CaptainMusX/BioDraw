/**
 * Settings Module
 * Manages sidebar icon display size preferences
 */

const STORAGE_KEYS = {
  ICON_SIZE: 'biodraw.iconDisplaySize'
};

const DEFAULTS = {
  iconSize: 80 // Default icon display size in pixels (80x80)
};

// Available preset sizes
const PRESET_SIZES = {
  small: 60,
  medium: 80,
  large: 120
};

/**
 * Initialize settings functionality
 */
export function initSettings() {
  // Listen for right-click on blank areas
  document.addEventListener('contextmenu', (e) => {
    // Ignore if clicking on interactive elements
    const target = e.target;
    if (target.closest('.icon-item') ||
        target.closest('button') ||
        target.closest('.category-menu') ||
        target.closest('input')) {
      return;
    }

    e.preventDefault();
    showSettingsDialog();
  });
}

/**
 * Show settings dialog
 */
function showSettingsDialog() {
  // Remove existing dialog if any
  const existing = document.getElementById('settings-dialog');
  if (existing) existing.remove();

  // Get current settings
  const currentSize = getIconDisplaySize();

  // Create dialog
  const dialog = document.createElement('div');
  dialog.id = 'settings-dialog';
  dialog.className = 'settings-dialog';
  dialog.innerHTML = `
    <div class="settings-content" role="dialog" aria-modal="true" aria-labelledby="settings-title">
      <h3 id="settings-title">Icon Display Settings</h3>
      <form id="settings-form">
        <div class="settings-section">
          <label class="settings-label">Default icon size (px):</label>
          <div class="preset-buttons">
            <button type="button" class="preset-btn ${currentSize === PRESET_SIZES.small ? 'active' : ''}" data-size="${PRESET_SIZES.small}">
              Small (60x60)
            </button>
            <button type="button" class="preset-btn ${currentSize === PRESET_SIZES.medium ? 'active' : ''}" data-size="${PRESET_SIZES.medium}">
              Medium (80x80)
            </button>
            <button type="button" class="preset-btn ${currentSize === PRESET_SIZES.large ? 'active' : ''}" data-size="${PRESET_SIZES.large}">
              Large (120x120)
            </button>
          </div>
        </div>
        <div class="settings-row">
          <label for="setting-icon-size">Custom size:</label>
          <input type="number" id="setting-icon-size" min="40" max="200" value="${currentSize}" required>
          <span class="settings-unit">px</span>
        </div>
        <div class="settings-buttons">
          <button type="submit" class="btn-primary">Save</button>
          <button type="button" class="btn-secondary" id="btn-cancel">Cancel</button>
        </div>
      </form>
    </div>
  `;

  document.body.appendChild(dialog);

  // Handle preset buttons
  const presetButtons = dialog.querySelectorAll('.preset-btn');
  const customInput = dialog.querySelector('#setting-icon-size');

  presetButtons.forEach((btn) => {
    btn.addEventListener('click', (event) => {
      event.preventDefault();
      const size = parseInt(btn.dataset.size, 10);
      customInput.value = size;

      // Update active state
      presetButtons.forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
    });
  });

  // Handle custom input change
  customInput.addEventListener('input', () => {
    // Remove active state from presets when user types custom value
    presetButtons.forEach((b) => b.classList.remove('active'));
  });

  // Handle form submission
  const form = dialog.querySelector('#settings-form');
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    saveSettings(form);
    dialog.remove();
  });

  // Handle cancel
  dialog.querySelector('#btn-cancel').addEventListener('click', () => {
    dialog.remove();
  });

  // Close on Esc
  const escHandler = (event) => {
    if (event.key === 'Escape') {
      dialog.remove();
      document.removeEventListener('keydown', escHandler);
    }
  };
  document.addEventListener('keydown', escHandler);

  // Close on outside click
  dialog.addEventListener('click', (event) => {
    if (event.target === dialog) {
      dialog.remove();
      document.removeEventListener('keydown', escHandler);
    }
  });

  // Focus first input
  customInput.focus();
  customInput.select();
}

/**
 * Save settings to localStorage and apply immediately
 */
function saveSettings(form) {
  const iconSize = parseInt(form.querySelector('#setting-icon-size').value, 10);

  // Validate
  if (isNaN(iconSize) || iconSize < 40 || iconSize > 200) {
    alert('Please enter a value between 40 and 200 pixels.');
    return;
  }

  // Save to localStorage
  localStorage.setItem(STORAGE_KEYS.ICON_SIZE, iconSize.toString());

  console.log('[BioDraw G3] Icon display size saved:', iconSize);
  
  // Apply immediately
  applyIconDisplaySize(iconSize);
  
  // Show confirmation
  showToast('Settings saved and applied.');
}

/**
 * Get icon display size from localStorage
 */
export function getIconDisplaySize() {
  const saved = localStorage.getItem(STORAGE_KEYS.ICON_SIZE);
  const size = saved ? parseInt(saved, 10) : DEFAULTS.iconSize;
  
  // Validate
  if (isNaN(size) || size < 40 || size > 200) {
    return DEFAULTS.iconSize;
  }
  
  return size;
}

/**
 * Apply icon display size to all icon items
 */
export function applyIconDisplaySize(size) {
  const targetSize = size || getIconDisplaySize();
  const gridElement = document.getElementById('icon-grid');
  const isListView = gridElement?.classList.contains('view-list');

  const iconImages = document.querySelectorAll('.icon-item img');

  iconImages.forEach((node) => {
    const img = node;
    if (!(img instanceof HTMLImageElement)) {
      return;
    }
    // Remove legacy width/height attributes so CSS can drive sizing
    img.removeAttribute('width');
    img.removeAttribute('height');

    // Reset inline style before applying new rules
    img.style.width = '';
    img.style.height = '';
    img.style.maxWidth = '';
    img.style.maxHeight = '';

    if (isListView) {
      img.style.width = 'auto';
      img.style.height = 'auto';
      img.style.maxHeight = `${targetSize}px`;
      img.style.maxWidth = `min(100%, ${Math.round(targetSize * 3)}px)`;
    } else {
      img.style.width = `${targetSize}px`;
      img.style.height = `${targetSize}px`;
      img.style.maxWidth = `${targetSize}px`;
      img.style.maxHeight = `${targetSize}px`;
    }
  });

  document.documentElement.style.setProperty('--icon-display-size', `${targetSize}px`);

  console.log(
    `[BioDraw Settings] Icon display size applied (${isListView ? 'list' : 'grid'} view): ` +
      `${targetSize}px to ${iconImages.length} icons`
  );
}

/**
 * Show a toast notification
 */
function showToast(message) {
  const toast = document.createElement('div');
  toast.className = 'toast-notification';
  toast.textContent = message;
  document.body.appendChild(toast);

  setTimeout(() => {
    toast.classList.add('show');
  }, 10);

  setTimeout(() => {
    toast.classList.remove('show');
    setTimeout(() => toast.remove(), 300);
  }, 2000);
}
