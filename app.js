/**
 * Freightoscope Bulk Import Leads Prototype Engine
 * Supports full Excel upload, parsing via SheetJS, 2-sheet template generation,
 * field and duplicate validation, permission toggling, and realistic UI state management.
 */

// Master Data Definitions
const COUNTRY_MASTER = {
  'US': 'United States',
  'IN': 'India',
  'AE': 'United Arab Emirates',
  'GB': 'United Kingdom',
  'SG': 'Singapore',
  'DE': 'Germany',
  'FR': 'France',
  'CN': 'China',
  'JP': 'Japan',
  'AU': 'Australia',
  'CA': 'Canada',
  'NL': 'Netherlands',
  'BR': 'Brazil',
  'IT': 'Italy',
  'ES': 'Spain',
  'MX': 'Mexico',
  'KR': 'South Korea',
  'SA': 'Saudi Arabia',
  'ZA': 'South Africa',
  'TR': 'Turkey',
  'DZ': 'Algeria',
  'MY': 'Malaysia',
  'VN': 'Vietnam',
  'TH': 'Thailand',
  'ID': 'Indonesia'
};

const USER_MASTER = [
  'sakshi.barnwal@freightoscope.com',
  'rafia.khan@freightoscope.com',
  'admin@freightoscope.com',
  'service@freightoscope.com',
  'john.doe@freightoscope.com',
  'anshul.p@freightoscope.com'
];

const LEAD_STAGES = [
  'New',
  'Contacted',
  'Qualified',
  'Proposal Sent',
  'Negotiation',
  'Won',
  'Lost'
];

// The 28 Predefined Column Headers
const DEFINED_HEADERS = [
  'Company Name*',
  'Short Name',
  'Legal Name',
  'Target Party Role',
  'Country Code*',
  'State',
  'City*',
  'Zip Code',
  'Address Line 1*',
  'Address Line 2',
  'Address Line 3',
  'Company Email',
  'Company Contact No.',
  'Company Website',
  'Source',
  'Lead Owner Email Id*',
  'Lead Stage*',
  'Industry',
  'Lead Temperature',
  'Service Interest',
  'Networks',
  'Notes',
  'Primary Contact',
  'First Name',
  'Last Name',
  'Email',
  'Phone No. (Primary)',
  'Designation'
];

// Initial Dummy Leads in Listing (Exact match to Screenshot 1)
let leadsData = [
  {
    leadId: 'RGBEXP',
    companyName: 'RGB Exports',
    location: 'New South Wales, Australia',
    stage: 'New',
    leadOwner: 'Ahmed Raza',
    temperature: '❄️',
    contactName: 'Mr. John Abel',
    contactPhone: '+61 412 345 678'
  },
  {
    leadId: 'GFA2026',
    companyName: 'GREAGR',
    location: 'Newyork, USA',
    stage: 'Contacted',
    leadOwner: 'Emily Carter',
    temperature: '🔥',
    contactName: 'Mrs. Britty Bell',
    contactPhone: '+1 412-555-0199'
  },
  {
    leadId: 'SBS2026',
    companyName: 'SkyBridge Solutions',
    location: 'Newyork, USA',
    stage: 'New',
    leadOwner: 'Ramesh Kumar',
    temperature: '❄️',
    contactName: 'Mr. Lukas Weber',
    contactPhone: '+1 412-555-0145'
  },
  {
    leadId: 'QES2026',
    companyName: 'QuantumEdge Systems',
    location: 'Newyork, USA',
    stage: 'Proposal Sent',
    leadOwner: 'Oliver Smith',
    temperature: '❄️',
    contactName: 'Mrs. Britty Bell',
    contactPhone: '+1 412-555-0182'
  },
  {
    leadId: 'BWL2026',
    companyName: 'BlueWave Logistics',
    location: 'Newyork, USA',
    stage: '--',
    leadOwner: 'Lukas Weber',
    temperature: '☀️',
    contactName: 'Mr. Bob Abel',
    contactPhone: '+1 412-555-0123'
  },
  {
    leadId: 'NTX2026',
    companyName: 'Nexora Technologies',
    location: 'Newyork, USA',
    stage: '--',
    leadOwner: 'Sarah Mitchell',
    temperature: '❄️',
    contactName: 'Mr. Bob Smith',
    contactPhone: '+1 412-555-0176'
  },
  {
    leadId: 'UNP2026',
    companyName: 'UrbanNest Properties',
    location: 'Newyork, USA',
    stage: '--',
    leadOwner: 'Daniel Tan',
    temperature: '🔥',
    contactName: 'Mr. Bob Smith',
    contactPhone: '+1 412-555-0194'
  },
  {
    leadId: 'ESI2026',
    companyName: 'EcoSphere Innovations',
    location: 'Newyork, USA',
    stage: '--',
    leadOwner: 'Laura Bennett',
    temperature: '❄️',
    contactName: 'Mr. Khalid Hassan',
    contactPhone: '+1 412-555-0131'
  },
  {
    leadId: 'FINA2026',
    companyName: 'FinCore Analytics',
    location: 'Newyork, USA',
    stage: 'New',
    leadOwner: 'Khalid Hassan',
    temperature: '☀️',
    contactName: 'Mrs. Anjali Nair',
    contactPhone: '+1 412-555-0158'
  }
];

// State variables
let importPermissionEnabled = true;
let simulateBackendError = false;
let currentUploadedFile = null;
let parsedLeadRecords = [];
let validationErrors = [];
let isProcessing = false;

// DOM Elements
const drawerBackdrop = document.getElementById('drawerBackdrop');
const importDrawer = document.getElementById('importDrawer');
const btnImport = document.getElementById('btnImport');
const btnCloseDrawer = document.getElementById('btnCloseDrawer');
const btnDrawerCancel = document.getElementById('btnDrawerCancel');
const btnDrawerReset = document.getElementById('btnDrawerReset');
const btnDrawerSave = document.getElementById('btnDrawerSave');
const downloadTemplateLink = document.getElementById('downloadTemplateLink');

const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const btnBrowseFiles = document.getElementById('btnBrowseFiles');
const uploadedFileBar = document.getElementById('uploadedFileBar');
const uploadedFileName = document.getElementById('uploadedFileName');
const uploadedFileSize = document.getElementById('uploadedFileSize');
const btnRemoveFile = document.getElementById('btnRemoveFile');

const bannerError = document.getElementById('bannerError');
const bannerErrorText = document.getElementById('bannerErrorText');
const errorsSection = document.getElementById('errorsSection');
const errorsTableBody = document.getElementById('errorsTableBody');
const errorCountPill = document.getElementById('errorCountPill');
const successValidationBox = document.getElementById('successValidationBox');
const successValidationText = document.getElementById('successValidationText');
const processingIndicator = document.getElementById('processingIndicator');

const leadsTableBody = document.getElementById('leadsTableBody');
const leadCountDisplay = document.getElementById('leadCountDisplay');
const importBtnWrapper = document.getElementById('importBtnWrapper');
const importTooltip = document.getElementById('importTooltip');

// Admin and Departments & Roles Submenu Elements
const menuAdmin = document.getElementById('menuAdmin');
const arrowAdmin = document.getElementById('arrowAdmin');
const submenuAdmin = document.getElementById('submenuAdmin');
const subitemDeptRoles = document.getElementById('subitemDeptRoles');

// Edit Roles & Permission Modal Elements
const rolesModalBackdrop = document.getElementById('rolesModalBackdrop');
const rolesModal = document.getElementById('rolesModal');
const btnCloseRolesModal = document.getElementById('btnCloseRolesModal');
const btnCancelRolesModal = document.getElementById('btnCancelRolesModal');
const btnSaveRolesModal = document.getElementById('btnSaveRolesModal');
const btnPermImport = document.getElementById('btnPermImport');
const wrapAllScreens = document.getElementById('wrapAllScreens');
const boxAllScreens = document.getElementById('boxAllScreens');
const searchScreensInput = document.getElementById('searchScreensInput');

// Sync Import Permission state across UI and button
function updateImportPermissionUI() {
  const btnPerm = document.getElementById('btnPermImport');
  if (btnPerm) {
    if (importPermissionEnabled) {
      btnPerm.classList.add('checked');
    } else {
      btnPerm.classList.remove('checked');
    }
  }

  if (importPermissionEnabled) {
    if (importBtnWrapper) importBtnWrapper.classList.remove('disabled');
    if (btnImport) {
      btnImport.classList.remove('disabled');
      btnImport.removeAttribute('disabled');
    }
  } else {
    if (importBtnWrapper) importBtnWrapper.classList.add('disabled');
    if (btnImport) {
      btnImport.classList.add('disabled');
      btnImport.setAttribute('disabled', 'true');
    }
  }
}

function toggleImportPermission(showToastNotice = false) {
  importPermissionEnabled = !importPermissionEnabled;
  updateImportPermissionUI();

  if (showToastNotice) {
    showToast(
      importPermissionEnabled ? 'Permission Granted' : 'Permission Revoked',
      importPermissionEnabled
        ? 'Import permission is ENABLED for Leads.'
        : 'Import permission is DISABLED for Leads.',
      importPermissionEnabled ? 'success' : 'warning',
      2500
    );
  }
}

function openRolesModal() {
  if (rolesModal && rolesModalBackdrop) {
    rolesModal.classList.add('open');
    rolesModalBackdrop.classList.add('open');
    updateImportPermissionUI();
  }
}

function closeRolesModal() {
  if (rolesModal && rolesModalBackdrop) {
    rolesModal.classList.remove('open');
    rolesModalBackdrop.classList.remove('open');
  }
}

// Helper to normalize header string: removes asterisks and dots, collapses spaces, trims
function cleanHeader(h) {
  if (!h) return '';
  return String(h)
    .replace(/[*_]/g, '')     // remove asterisks/underscores first
    .replace(/\s+/g, ' ')     // normalize multiple spaces to single space
    .trim()                   // trim leading and trailing spaces AFTER removing symbols!
    .toLowerCase();           // lowercase for case-insensitive matching
}

// Map clean headers to standard keys
const HEADER_KEY_MAP = {
  'company name': 'companyName',
  'short name': 'shortName',
  'legal name': 'legalName',
  'target party role': 'targetPartyRole',
  'country code': 'countryCode',
  'state': 'state',
  'city': 'city',
  'zip code': 'zipCode',
  'address line 1': 'addressLine1',
  'address line 2': 'addressLine2',
  'address line 3': 'addressLine3',
  'company email': 'companyEmail',
  'company contact no.': 'companyContactNo',
  'company contact no': 'companyContactNo',
  'company website': 'companyWebsite',
  'source': 'source',
  'lead owner email id': 'leadOwnerEmailId',
  'lead stage': 'leadStage',
  'industry': 'industry',
  'lead temperature': 'leadTemperature',
  'service interest': 'serviceInterest',
  'networks': 'networks',
  'notes': 'notes',
  'primary contact': 'primaryContact',
  'first name': 'firstName',
  'last name': 'lastName',
  'email': 'email',
  'phone no. (primary)': 'phonePrimary',
  'phone no (primary)': 'phonePrimary',
  'phone no.': 'phonePrimary',
  'designation': 'designation'
};

// Initial Render
document.addEventListener('DOMContentLoaded', () => {
  renderLeadsTable();
  updateImportPermissionUI();
  setupEventListeners();
});

function setupEventListeners() {
  // Admin Submenu Toggle
  if (menuAdmin && submenuAdmin) {
    menuAdmin.addEventListener('click', () => {
      const isCurrentlyOpen = submenuAdmin.style.display === 'block';
      submenuAdmin.style.display = isCurrentlyOpen ? 'none' : 'block';
      if (arrowAdmin) {
        arrowAdmin.style.transform = isCurrentlyOpen ? 'rotate(0deg)' : 'rotate(90deg)';
      }
    });
  }

  // Departments & Roles Submenu Item Click -> Opens Edit Roles & Permission Modal
  if (subitemDeptRoles) {
    subitemDeptRoles.addEventListener('click', (e) => {
      e.stopPropagation();
      openRolesModal();
    });
  }

  // Modal Close, Cancel, and Backdrop Handlers
  if (btnCloseRolesModal) btnCloseRolesModal.addEventListener('click', closeRolesModal);
  if (btnCancelRolesModal) btnCancelRolesModal.addEventListener('click', closeRolesModal);
  if (rolesModalBackdrop) rolesModalBackdrop.addEventListener('click', closeRolesModal);

  // Modal Import Permission Button Toggle (Clean, single-click toggle with zero double-trigger)
  if (btnPermImport) {
    btnPermImport.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      toggleImportPermission(true);
    });
  }

  // Modal Save Button
  if (btnSaveRolesModal) {
    btnSaveRolesModal.addEventListener('click', () => {
      updateImportPermissionUI();
      closeRolesModal();
      showToast(
        importPermissionEnabled ? 'Permissions Saved' : 'Permissions Updated',
        importPermissionEnabled
          ? 'Import permission is ENABLED. You can now import leads.'
          : 'Import permission is DISABLED. Import button is now disabled.',
        importPermissionEnabled ? 'success' : 'warning',
        4000
      );
    });
  }

  // Module items click selection
  document.querySelectorAll('.module-item').forEach(item => {
    item.addEventListener('click', () => {
      document.querySelectorAll('.module-item').forEach(m => m.classList.remove('active'));
      item.classList.add('active');
    });
  });

  // Action pills toggle (Add, Delete, Update, View)
  document.querySelectorAll('.btn-action-pill').forEach(pill => {
    if (pill.id === 'btnPermImport') return; // Handled exclusively above
    pill.addEventListener('click', (e) => {
      e.preventDefault();
      pill.classList.toggle('checked');
    });
  });

  // Screen row left checkbox toggle
  document.querySelectorAll('.screen-row-left').forEach(rowLeft => {
    rowLeft.addEventListener('click', (e) => {
      e.preventDefault();
      const box = rowLeft.querySelector('.screen-check-box');
      if (box) {
        box.classList.toggle('checked');
        const isChecked = box.classList.contains('checked');
        const row = rowLeft.closest('.screen-row');
        if (row) {
          row.querySelectorAll('.btn-action-pill').forEach(pill => {
            pill.classList.toggle('checked', isChecked);
          });
          if (row.id === 'rowLeadsScreen') {
            importPermissionEnabled = isChecked;
            updateImportPermissionUI();
          }
        }
      }
    });
  });

  // All Screens Toggle
  if (wrapAllScreens && boxAllScreens) {
    wrapAllScreens.addEventListener('click', (e) => {
      e.preventDefault();
      boxAllScreens.classList.toggle('checked');
      const allChecked = boxAllScreens.classList.contains('checked');
      document.querySelectorAll('.screen-check-box').forEach(b => {
        b.classList.toggle('checked', allChecked);
      });
      document.querySelectorAll('.btn-action-pill').forEach(p => {
        p.classList.toggle('checked', allChecked);
      });
      importPermissionEnabled = allChecked;
      updateImportPermissionUI();
    });
  }

  // Search screens input
  if (searchScreensInput) {
    searchScreensInput.addEventListener('input', (e) => {
      const q = e.target.value.toLowerCase().trim();
      document.querySelectorAll('.screens-list-container .screen-row').forEach(row => {
        const text = row.textContent.toLowerCase();
        row.style.display = text.includes(q) ? 'flex' : 'none';
      });
    });
  }


  // Import Button & Wrapper Click (Enforces Permission & Tooltip)
  if (importBtnWrapper) {
    importBtnWrapper.addEventListener('click', (e) => {
      if (!importPermissionEnabled) {
        // Flash tooltip
        importBtnWrapper.classList.add('show-tooltip');
        setTimeout(() => importBtnWrapper.classList.remove('show-tooltip'), 3000);

        showToast(
          'Access Restricted',
          'You do not have permission to perform this action. Please contact your Admin or service@freightoscope.com for assistance',
          'error',
          6500
        );
        return;
      }
      openDrawer();
    });
  }

  if (btnImport) {
    btnImport.addEventListener('click', (e) => {
      if (!importPermissionEnabled) {
        e.preventDefault();
        e.stopPropagation();
        importBtnWrapper.classList.add('show-tooltip');
        setTimeout(() => importBtnWrapper.classList.remove('show-tooltip'), 3000);
        showToast(
          'Access Restricted',
          'You do not have permission to perform this action. Please contact your Admin or service@freightoscope.com for assistance',
          'error',
          6500
        );
        return;
      }
      openDrawer();
    });
  }

  // Close Drawer
  btnCloseDrawer.addEventListener('click', closeDrawer);
  btnDrawerCancel.addEventListener('click', closeDrawer);
  drawerBackdrop.addEventListener('click', closeDrawer);

  // Download Template (generates 2 sheets)
  downloadTemplateLink.addEventListener('click', (e) => {
    e.preventDefault();
    generateAndDownloadTemplate();
  });

  // Browse files button
  btnBrowseFiles.addEventListener('click', (e) => {
    e.stopPropagation();
    fileInput.click();
  });

  dropzone.addEventListener('click', () => {
    fileInput.click();
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) {
      handleSelectedFile(e.target.files[0]);
    }
  });

  // Drag and Drop
  ['dragenter', 'dragover'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add('drag-over');
    }, false);
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove('drag-over');
    }, false);
  });

  dropzone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    if (dt && dt.files && dt.files[0]) {
      handleSelectedFile(dt.files[0]);
    }
  });

  // Remove File
  btnRemoveFile.addEventListener('click', (e) => {
    e.stopPropagation();
    resetUploadSection();
  });

  // Reset Button
  btnDrawerReset.addEventListener('click', () => {
    resetUploadSection();
  });

  // Save Button
  btnDrawerSave.addEventListener('click', handleFinalSave);
}

// Drawer visibility controls
function openDrawer() {
  importDrawer.classList.add('open');
  drawerBackdrop.classList.add('open');
}

function closeDrawer() {
  if (isProcessing) {
    // Abort processing as specified: "If a file is currently Processing, clicking close must abort the operation."
    isProcessing = false;
    hideProcessing();
  }
  importDrawer.classList.remove('open');
  drawerBackdrop.classList.remove('open');
}

// Reset upload section
function resetUploadSection() {
  currentUploadedFile = null;
  fileInput.value = '';
  parsedLeadRecords = [];
  validationErrors = [];
  isProcessing = false;

  uploadedFileBar.style.display = 'none';
  dropzone.style.display = 'flex';
  btnDrawerReset.style.display = 'none';
  hideBanners();
  hideErrors();
  hideSuccessBox();
  hideProcessing();

  setSaveEnabled(false);
}

function setSaveEnabled(enabled) {
  if (enabled) {
    btnDrawerSave.disabled = false;
    btnDrawerSave.classList.add('active');
  } else {
    btnDrawerSave.disabled = true;
    btnDrawerSave.classList.remove('active');
  }
}

function showProcessing() {
  isProcessing = true;
  processingIndicator.style.display = 'flex';
}

function hideProcessing() {
  isProcessing = false;
  processingIndicator.style.display = 'none';
}

function showBannerError(msg) {
  bannerErrorText.textContent = msg;
  bannerError.style.display = 'block';
}

function hideBanners() {
  bannerError.style.display = 'none';
  bannerErrorText.textContent = '';
}

function hideErrors() {
  errorsSection.style.display = 'none';
  errorsTableBody.innerHTML = '';
}

function hideSuccessBox() {
  successValidationBox.style.display = 'none';
  successValidationText.textContent = '';
}

// File Validation & Handling
function handleSelectedFile(file) {
  hideBanners();
  hideErrors();
  hideSuccessBox();
  setSaveEnabled(false);

  const fileName = file.name;
  const ext = fileName.slice((fileName.lastIndexOf(".") - 1 >>> 0) + 2).toLowerCase();

  // 1. File Format Restriction: Accept only .xls and .xlsx
  if (ext !== 'xlsx' && ext !== 'xls') {
    // Other formats are rejected, not displayed, and show: "Invalid format found. Please Select a valid format"
    showBannerError('Invalid format found. Please Select a valid format');
    fileInput.value = '';
    return;
  }

  // Display file in blue bar below upload area
  currentUploadedFile = file;
  uploadedFileName.textContent = file.name;
  uploadedFileSize.textContent = formatBytes(file.size);
  uploadedFileBar.style.display = 'flex';
  btnDrawerReset.style.display = 'inline-block';

  showProcessing();

  // Parse Excel via SheetJS
  const reader = new FileReader();
  reader.onload = (e) => {
    if (!isProcessing) return; // aborted
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      processWorkbook(workbook);
    } catch (err) {
      hideProcessing();
      showBannerError('Failed to parse the file. Please ensure it is a valid Excel spreadsheet.');
    }
  };
  reader.onerror = () => {
    hideProcessing();
    showBannerError('Error reading file. Please try again.');
  };
  reader.readAsArrayBuffer(file);
}

function formatBytes(bytes, decimals = 1) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const dm = decimals < 0 ? 0 : decimals;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

// Process Workbook and Validate
function processWorkbook(workbook) {
  hideProcessing();

  // Find sheet: look for sheet containing lead data, ignoring instructions
  let sheetName = workbook.SheetNames.find(n =>
    (n.toLowerCase().includes('bulk') || n.toLowerCase().includes('lead') || n.toLowerCase().includes('import')) &&
    !n.toLowerCase().includes('instruction')
  );

  if (!sheetName) {
    for (const name of workbook.SheetNames) {
      if (name.toLowerCase().includes('instruction')) continue;
      sheetName = name;
      break;
    }
  }

  if (!sheetName) {
    sheetName = workbook.SheetNames[0];
  }

  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    showBannerError('No valid sheet found in uploaded Excel file.');
    return;
  }

  // Read as array of arrays
  const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  if (!rawData || rawData.length === 0) {
    showBannerError('No lead found in the uploaded file. Please add at least one lead detail and upload again.');
    return;
  }

  // Header row (Row 1)
  const headerRow = rawData[0];
  if (!headerRow || headerRow.length === 0) {
    showBannerError('Incorrect column names found. Please upload a correct excel file');
    return;
  }

  // 2. File Parsing & Header Validation
  // Case-insensitively matched, column order irrelevant.
  // Missing columns shall block import and display: "Incorrect column names found. Please upload a correct excel file"
  // For duplicate columns, only first occurrence considered, subsequent ignored and not overwrite.

  const columnMap = {}; // cleanHeader -> first col index
  const foundCleanHeaders = new Set();

  headerRow.forEach((colName, colIdx) => {
    const clean = cleanHeader(colName);
    if (clean && !foundCleanHeaders.has(clean)) {
      foundCleanHeaders.add(clean);
      columnMap[clean] = colIdx;
    }
  });

  // Verify all 28 defined headers are present
  let missingRequiredTemplateHeaders = false;
  for (const defHeader of DEFINED_HEADERS) {
    const clean = cleanHeader(defHeader);
    if (!foundCleanHeaders.has(clean)) {
      missingRequiredTemplateHeaders = true;
      break;
    }
  }

  if (missingRequiredTemplateHeaders) {
    showBannerError('Incorrect column names found. Please upload a correct excel file');
    return;
  }

  // 3. Extract Rows & Empty file check
  const rows = rawData.slice(1);
  const records = [];

  rows.forEach((rowArr, index) => {
    const rowNum = index + 2; // Excel row number (1-based, Row 1 is header)
    
    // Check if entire row is empty
    const isRowEmpty = rowArr.every(val => val === '' || val === null || val === undefined || String(val).trim() === '');
    if (!isRowEmpty) {
      const record = { _excelRow: rowNum };
      for (const defHeader of DEFINED_HEADERS) {
        const clean = cleanHeader(defHeader);
        const colIdx = columnMap[clean];
        const val = (colIdx !== undefined && rowArr[colIdx] !== undefined) ? String(rowArr[colIdx]).trim() : '';
        const key = HEADER_KEY_MAP[clean] || clean;
        record[key] = val;
      }
      records.push(record);
    }
  });

  // "If the uploaded Excel file contains no Lead details and all Lead fields are blank, the system shall display:
  // 'No lead found in the uploaded file. Please add at least one lead detail and upload again.'"
  if (records.length === 0) {
    showBannerError('No lead found in the uploaded file. Please add at least one lead detail and upload again.');
    return;
  }

  parsedLeadRecords = records;

  // 4. Run Field Validations and Duplicate Validations
  runValidationSuite(records);
}

// Check if an email belongs to FOS (Freightoscope internal domain)
function isFosEmail(email) {
  if (!email || typeof email !== 'string') return false;
  const trimmed = email.trim().toLowerCase();
  const atIdx = trimmed.lastIndexOf('@');
  if (atIdx === -1 || atIdx === 0 || atIdx === trimmed.length - 1) return false;
  const domain = trimmed.slice(atIdx + 1);
  if (!domain) return false;

  // Exact FOS domains or subdomains (*.fos.com, *.freightoscope.com)
  if (domain === 'fos.com' || domain.endsWith('.fos.com')) return true;
  if (domain === 'freightoscope.com' || domain.endsWith('.freightoscope.com')) return true;

  // Also support any regional FOS domain (e.g. fos.in, freightoscope.ae)
  if (/^([a-zA-Z0-9-]+\.)*(fos|freightoscope)\.[a-zA-Z]{2,}$/i.test(domain)) return true;

  return false;
}

// Validate company website URL format
function isValidCompanyWebsite(urlStr) {
  if (!urlStr || typeof urlStr !== 'string') return false;
  const str = urlStr.trim();
  if (!str) return false;

  // Reject if contains whitespace
  if (/\s/.test(str)) return false;

  // Reject typos of protocol like "https." or "http." or "http:/" or "https:/" or "http//"
  if (/^https?[\.:\/]/i.test(str) && !/^https?:\/\//i.test(str)) {
    return false;
  }

  // Reject other protocols (ftp:, mailto:, file:, javascript:)
  if (/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(str) && !/^https?:\/\//i.test(str)) {
    return false;
  }

  // Strip http:// or https:// for domain parsing
  let withoutProtocol = str.replace(/^https?:\/\//i, '');

  // Separate path/query/hash from host
  const slashIdx = withoutProtocol.indexOf('/');
  const questionIdx = withoutProtocol.indexOf('?');
  const hashIdx = withoutProtocol.indexOf('#');
  let cutIdx = withoutProtocol.length;
  if (slashIdx !== -1 && slashIdx < cutIdx) cutIdx = slashIdx;
  if (questionIdx !== -1 && questionIdx < cutIdx) cutIdx = questionIdx;
  if (hashIdx !== -1 && hashIdx < cutIdx) cutIdx = hashIdx;

  const hostPort = withoutProtocol.slice(0, cutIdx);
  const host = hostPort.split(':')[0]; // strip optional port

  if (!host || !host.includes('.')) return false;
  if (host.startsWith('.') || host.endsWith('.') || host.startsWith('-') || host.endsWith('-')) return false;

  const parts = host.split('.');
  if (parts.length < 2) return false;

  // Subdomain cannot be "http" or "https" (e.g. "https.dbjbf.com")
  if (parts[0].toLowerCase() === 'http' || parts[0].toLowerCase() === 'https') {
    return false;
  }

  // Verify each label
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    if (!part || part.length === 0) return false;
    if (!/^[a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?$/.test(part)) {
      return false;
    }
  }

  // TLD must be at least 2 alphabetical characters
  const tld = parts[parts.length - 1];
  if (!/^[a-zA-Z]{2,}$/.test(tld)) {
    return false;
  }

  return true;
}

// Validation Suite
function runValidationSuite(records) {
  validationErrors = [];

  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const iso2Regex = /^[a-zA-Z]{2}$/;
  const phoneLetterRegex = /[a-zA-Z]/;

  // Track duplicates: Company Name + City + Country (normalized)
  const duplicateLeadMap = new Map();

  records.forEach((rec, idx) => {
    const row = rec._excelRow;

    // Company Name
    // Empty -> "Company Name is missing"
    // > 100 chars -> "Company Name cannot exceed 100 characters"
    if (!rec.companyName) {
      addError(row, 'Company Name', 'Missing Field', 'Company Name is missing');
    } else if (rec.companyName.length > 100) {
      addError(row, 'Company Name', 'Character Limit Exceeded', 'Company Name cannot exceed 100 characters');
    }

    // Country Code
    // Empty -> "Country Code is missing"
    // Not valid 2-letter ISO code -> "Country Code must be a valid 2-letter ISO code"
    // Does not match existing Country -> "This country could not be found"
    if (!rec.countryCode) {
      addError(row, 'Country Code', 'Missing Field', 'Country Code is missing');
    } else if (!iso2Regex.test(rec.countryCode)) {
      addError(row, 'Country Code', 'Invalid Format', 'Country Code must be a valid 2-letter ISO code');
    } else if (!COUNTRY_MASTER[rec.countryCode.toUpperCase()]) {
      addError(row, 'Country Code', 'Country Not Found', 'This country could not be found');
    }

    // City
    // Empty -> "City is missing"
    // > 100 chars -> "City cannot exceed 100 characters"
    if (!rec.city) {
      addError(row, 'City', 'Missing Field', 'City is missing');
    } else if (rec.city.length > 100) {
      addError(row, 'City', 'Character Limit Exceeded', 'City cannot exceed 100 characters');
    }

    // Address Line 1
    // Empty -> "Address Line 1 is missing"
    // > 500 chars -> "Address Line 1 cannot exceed 500 characters"
    if (!rec.addressLine1) {
      addError(row, 'Address Line 1', 'Missing Field', 'Address Line 1 is missing');
    } else if (rec.addressLine1.length > 500) {
      addError(row, 'Address Line 1', 'Character Limit Exceeded', 'Address Line 1 cannot exceed 500 characters');
    }

    // Company Email
    if (rec.companyEmail) {
      const cEmail = rec.companyEmail.trim();
      if (!emailRegex.test(cEmail)) {
        addError(row, 'Company Email', 'Invalid Format', 'This company email address is not valid');
      } else if (cEmail.length > 100) {
        addError(row, 'Company Email', 'Character Limit Exceeded', 'Company email cannot exceed 100 characters');
      }
    }

    // Company Contact No.
    if (rec.companyContactNo) {
      const cPhone = rec.companyContactNo.trim();
      if (cPhone.length > 20) {
        addError(row, 'Company Contact No.', 'Character Limit Exceeded', 'Company contact number cannot exceed 20 characters');
      } else if (phoneLetterRegex.test(cPhone)) {
        addError(row, 'Company Contact No.', 'Invalid Format', 'Company contact number can only contain digits & special characters');
      }
    }

    // Company Website
    // If entered:
    // > 100 chars -> "Company website cannot exceed 100 characters"
    // Invalid format -> "This company website is not valid"
    if (rec.companyWebsite) {
      const site = rec.companyWebsite.trim();
      if (site.length > 100) {
        addError(row, 'Company Website', 'Character Limit Exceeded', 'Company website cannot exceed 100 characters');
      } else if (!isValidCompanyWebsite(site)) {
        addError(row, 'Company Website', 'Invalid Format', 'This company website is not valid');
      }
    }

    // Lead Owner Email Id
    // Empty -> "Lead Owner Email Id is missing"
    // > 100 chars -> "Lead Owner Email Id cannot exceed 100 characters"
    // Invalid email format -> "This owner email address is not valid"
    // Not an FOS mail -> "This lead owner could not be found"
    const ownerEmail = (rec.leadOwnerEmailId || '').trim();
    if (!ownerEmail) {
      addError(row, 'Lead Owner Email Id', 'Missing Field', 'Lead Owner Email Id is missing');
    } else if (ownerEmail.length > 100) {
      addError(row, 'Lead Owner Email Id', 'Character Limit Exceeded', 'Lead Owner Email Id cannot exceed 100 characters');
    } else if (!emailRegex.test(ownerEmail)) {
      addError(row, 'Lead Owner Email Id', 'Invalid Format', 'This owner email address is not valid');
    } else if (!isFosEmail(ownerEmail)) {
      addError(row, 'Lead Owner Email Id', 'Owner Not Found', 'This lead owner could not be found');
    }

    // Lead Stage
    // Empty -> "Lead Stage is missing"
    if (!rec.leadStage) {
      addError(row, 'Lead Stage', 'Missing Field', 'Lead Stage is missing');
    }

    // Primary Contact: must be TRUE, FALSE, or blank
    let primaryContactRaw = rec.primaryContact !== undefined && rec.primaryContact !== null ? String(rec.primaryContact).trim() : '';
    let isPrimaryTrue = false;
    let isPrimaryFalse = false;

    if (primaryContactRaw === '1' || primaryContactRaw.toUpperCase() === 'TRUE') {
      isPrimaryTrue = true;
    } else if (primaryContactRaw === '0' || primaryContactRaw.toUpperCase() === 'FALSE' || primaryContactRaw === '') {
      isPrimaryFalse = true;
    } else {
      addError(row, 'Primary Contact', 'Invalid Value', 'Primary Contact must be TRUE or FALSE');
    }

    // Primary Contact = TRUE validations
    if (isPrimaryTrue) {
      // First Name
      if (!rec.firstName) {
        addError(row, 'First Name', 'Missing Field', 'First Name is missing (required when Primary Contact is TRUE)');
      } else if (rec.firstName.length > 50) {
        addError(row, 'First Name', 'Character Limit Exceeded', 'First Name cannot exceed 50 characters');
      }

      // Last Name
      if (!rec.lastName) {
        addError(row, 'Last Name', 'Missing Field', 'Last Name is missing (required when Primary Contact is TRUE)');
      } else if (rec.lastName.length > 50) {
        addError(row, 'Last Name', 'Character Limit Exceeded', 'Last Name cannot exceed 50 characters');
      }

      // Email
      if (!rec.email) {
        addError(row, 'Email', 'Missing Field', 'Email is missing (required when Primary Contact is TRUE)');
      } else if (!emailRegex.test(rec.email)) {
        addError(row, 'Email', 'Invalid Format', 'This email address is not valid');
      } else if (rec.email.length > 100) {
        addError(row, 'Email', 'Character Limit Exceeded', 'Email cannot exceed 100 characters');
      }

      // Phone No. (Primary)
      if (rec.phonePrimary) {
        if (phoneLetterRegex.test(rec.phonePrimary)) {
          addError(row, 'Phone No. (Primary)', 'Invalid Format', 'Phone number can only contain digits & special characters');
        } else {
          // Count digits
          const digitCount = (rec.phonePrimary.match(/\d/g) || []).length;
          if (digitCount > 20) {
            addError(row, 'Phone No. (Primary)', 'Character Limit Exceeded', 'Phone number cannot exceed 20 digits');
          }
        }
      }
    }

    // Build Duplicate Key (Company Name + City + Country)
    const normComp = (rec.companyName || '').trim().toLowerCase();
    const normCity = (rec.city || '').trim().toLowerCase();
    const normCountry = (rec.countryCode || '').trim().toLowerCase();
    
    if (normComp && normCity && normCountry) {
      const dupKey = `${normComp}###${normCity}###${normCountry}`;
      if (!duplicateLeadMap.has(dupKey)) {
        duplicateLeadMap.set(dupKey, [row]);
      } else {
        duplicateLeadMap.get(dupKey).push(row);
      }
    }
  });

  // Check Duplicate Entries
  for (const [key, rowList] of duplicateLeadMap.entries()) {
    if (rowList.length > 1) {
      // Multiple rows share the same Company Name, City, Country
      rowList.forEach(rNum => {
        addError(
          rNum,
          'Company Name + City + Country',
          'Duplicate Entry',
          'This lead is already present in the uploaded file'
        );
      });
    }
  }

  // Display Results
  displayValidationSummary(records.length);
}

function addError(row, col, type, message) {
  validationErrors.push({ row, col, type, message });
}

function displayValidationSummary(totalRecords) {
  if (validationErrors.length > 0) {
    // Sort errors by Excel Row then Column
    validationErrors.sort((a, b) => a.row - b.row);

    errorCountPill.textContent = `${validationErrors.length} ${validationErrors.length === 1 ? 'Error' : 'Errors'}`;
    errorsTableBody.innerHTML = '';

    validationErrors.forEach(err => {
      const tr = document.createElement('tr');

      let badgeClass = 'format';
      if (err.type === 'Missing Field') badgeClass = 'missing';
      else if (err.type === 'Character Limit Exceeded') badgeClass = 'limit';
      else if (err.type === 'Duplicate Entry') badgeClass = 'duplicate';
      else if (err.type === 'Country Not Found' || err.type === 'Owner Not Found') badgeClass = 'notfound';
      else if (err.type === 'Invalid Value') badgeClass = 'invalid';

      tr.innerHTML = `
        <td style="font-weight: 700; color: #1e293b;">${err.row}</td>
        <td style="font-weight: 600; color: #334155;">${escapeHtml(err.col)}</td>
        <td><span class="error-badge ${badgeClass}">${escapeHtml(err.type)}</span></td>
        <td style="color: #b42318; font-weight: 500;">${escapeHtml(err.message)}</td>
      `;
      errorsTableBody.appendChild(tr);
    });

    errorsSection.style.display = 'flex';
    hideSuccessBox();
    setSaveEnabled(false);
  } else {
    // All rows passed validation
    hideErrors();
    successValidationText.textContent = `All rows passed validation. ${totalRecords} lead(s) are ready to import.`;
    successValidationBox.style.display = 'flex';
    setSaveEnabled(true);
  }
}

// Final Save Processing
function handleFinalSave() {
  if (validationErrors.length > 0 || parsedLeadRecords.length === 0) return;

  // Check Backend Error Simulation
  if (simulateBackendError) {
    showToast(
      'Import Failed',
      'Failed to import leads. Please try again',
      'error',
      6000
    );
    // As per requirement: "The drawer remains open"
    return;
  }

  // Generate unique Lead IDs and Map Primary Contacts
  const newLeads = parsedLeadRecords.map((rec, idx) => {
    // Lead ID Generation: prefix from company name + random/seq
    const cleanComp = rec.companyName.replace(/[^a-zA-Z]/g, '').toUpperCase().slice(0, 6) || 'LEAD';
    const seq = Math.floor(100 + Math.random() * 900);
    const generatedLeadId = `${cleanComp}${seq}`;

    // Country Display
    const countryName = COUNTRY_MASTER[(rec.countryCode || '').toUpperCase()] || rec.countryCode;
    const locationStr = rec.state ? `${rec.city}, ${rec.state}, ${countryName}` : `${rec.city}, ${countryName}`;

    // Primary Contact mapping
    let primaryContactDisplay = '--';
    const pcStr = (rec.primaryContact || '').toString().trim().toUpperCase();
    if (pcStr === '1' || pcStr === 'TRUE') {
      primaryContactDisplay = `${rec.firstName} ${rec.lastName}`.trim();
    }

    // Lead Owner display name
    let leadOwnerDisplay = rec.leadOwnerEmailId;
    if (rec.leadOwnerEmailId.includes('sakshi')) leadOwnerDisplay = 'Sakshi Barnwal';
    else if (rec.leadOwnerEmailId.includes('rafia')) leadOwnerDisplay = 'Rafia Khan';
    else if (rec.leadOwnerEmailId.includes('john')) leadOwnerDisplay = 'John Doe';
    else if (rec.leadOwnerEmailId.includes('anshul')) leadOwnerDisplay = 'Anshul P';
    else if (rec.leadOwnerEmailId.includes('admin')) leadOwnerDisplay = 'Admin';

    return {
      leadId: generatedLeadId,
      companyName: rec.companyName,
      location: locationStr,
      stage: rec.leadStage || 'New',
      temperature: rec.leadTemperature === 'Warm' ? '☀️' : (rec.leadTemperature === 'Cold' ? '❄️' : '🔥'),
      leadOwner: leadOwnerDisplay,
      contactName: primaryContactDisplay,
      contactPhone: rec.phonePrimary || '--',
      isNew: true
    };
  });

  // Prepend new leads to table
  leadsData = [...newLeads, ...leadsData];
  renderLeadsTable();

  // Close Drawer
  closeDrawer();
  resetUploadSection();

  // Success toast
  showToast(
    'Success',
    'Leads imported successfully',
    'success'
  );
}

// Render Leads Listing Table (Exact match to Screenshot 1)
function renderLeadsTable() {
  leadsTableBody.innerHTML = '';
  leadCountDisplay.textContent = leadsData.length;

  leadsData.forEach(lead => {
    const tr = document.createElement('tr');
    if (lead.isNew) tr.classList.add('newly-imported');

    let stageBadge = `<span class="badge-new-solid">${escapeHtml(lead.stage)}</span>`;
    if (lead.stage === 'Contacted') {
      stageBadge = `<span class="badge-contacted-solid">${escapeHtml(lead.stage)}</span>`;
    } else if (lead.stage === 'Proposal Sent') {
      stageBadge = `<span class="badge-proposal-solid">${escapeHtml(lead.stage)}</span>`;
    } else if (lead.stage === 'Qualified') {
      stageBadge = `<span class="badge-qualified-solid">${escapeHtml(lead.stage)}</span>`;
    } else if (lead.stage === '--') {
      stageBadge = `<span style="color: #9ca3af; font-weight: 500;">--</span>`;
    }

    tr.innerHTML = `
      <td style="width: 32px; text-align: center;"></td>
      <td class="icon-cell">⋮</td>
      <td class="icon-cell" style="color: #d1d5db;">☆</td>
      <td><a class="lead-id-link">${escapeHtml(lead.leadId)}</a></td>
      <td>
        <div class="lead-company-name">${escapeHtml(lead.companyName)}</div>
        <div class="lead-location-sub">${escapeHtml(lead.location)}</div>
      </td>
      <td>${stageBadge}</td>
      <td>${escapeHtml(lead.leadOwner)}</td>
      <td style="font-size: 14px; text-align: center;">${lead.temperature}</td>
      <td>${escapeHtml(lead.contactName || '--')}</td>
      <td>${escapeHtml(lead.contactPhone || '--')}</td>
    `;
    leadsTableBody.appendChild(tr);
  });
}

// Generate & Download Real 2-Sheet Excel Template
function generateAndDownloadTemplate() {
  if (typeof XLSX === 'undefined') {
    showToast('Error', 'SheetJS library not loaded. Please ensure xlsx.full.min.js is present.', 'error');
    return;
  }

  const wb = XLSX.utils.book_new();

  // 1. Sheet 1: "Import Bulk Lead"
  const sheet1Data = [
    DEFINED_HEADERS,
    [
      'Acme Global Logistics',
      'AGL',
      'Acme Global Logistics Pvt. Ltd.',
      'Customer',
      'US',
      'Texas',
      'Dallas',
      '75001',
      '1200 Logistics Blvd, Suite 400',
      'Building B',
      '',
      'info@acme-globallog.com',
      '+1 214-555-0199',
      'www.acme-globallog.com',
      'Direct / Inbound',
      'sakshi.barnwal@freightoscope.com',
      'New',
      'Logistics',
      'Hot',
      'Ocean Freight (FCL)',
      'WCA',
      'Referred from partner network. Interested in FCL contracts.',
      'TRUE',
      'Sharon',
      'Roy',
      'sharon.roy@acme-globallog.com',
      '+1 214-555-0145',
      'Logistics Manager'
    ]
  ];
  const ws1 = XLSX.utils.aoa_to_sheet(sheet1Data);

  // Auto-width for Sheet 1
  ws1['!cols'] = DEFINED_HEADERS.map(h => ({ wch: Math.max(h.length + 3, 16) }));

  // 2. Sheet 2: "Instructions"
  const instructionHeaders = [
    'Field',
    'Field Value',
    'No of Characters & Type',
    'Mandatory/Optional',
    'Other Validations'
  ];

  const instructionRows = [
    ['Company Name*', 'Free Text', '1-100 Alpha Numeric & Special Characters', 'Mandatory', '1. Company Name must have a value.\n2. Value must not exceed 100 characters.'],
    ['Short Name', 'Free Text', '1-100 Alpha Numeric & Special Characters', 'Optional', '1. Value must not exceed 100 characters.'],
    ['Legal Name', 'Free Text', '1-100 Alpha Numeric & Special Characters', 'Optional', '1. Value must not exceed 100 characters.'],
    ['Target Party Role', 'Dropdown (Hardcoded)', 'Value from Dropdown (Agent, Customer)', 'Optional', '1. Each cell of the column contains the Target Party Role options: Agent, Customer.\n2. One of the options should be selected from the dropdown.\n3. If no option is selected or an incorrect value is entered, Target Party Role shall be empty for the Lead.'],
    ['Country Code*', 'Lookup', '2 Alpha Characters (ISO Code)', 'Mandatory', '1. Country Code must have a value.\n2. Value must be a valid 2-letter ISO Country Code.\n3. Value must match an existing Country. Source: Master -> Locations -> Countries.'],
    ['State', 'Lookup (Free text fallback)', '1-100 Alpha Numeric & Special Characters', 'Optional', '1. If the value matches an existing State (under the selected Country) in Master -> Locations -> States, it is linked to that State.\n2. If the value does not match an existing State, the State field is imported blank for the Lead - no error is raised.'],
    ['City*', 'Free Text', '1-100 Alpha Numeric & Special Characters', 'Mandatory', '1. City must have a value.\n2. Value must not exceed 100 characters.\n3. Accepts free text; not validated against a master.'],
    ['Zip Code', 'Free Text', '1-20 Alpha Numeric & Special Characters', 'Optional', '1. Value must not exceed 20 characters.'],
    ['Address Line 1*', 'Free Text', '1-500 Alpha Numeric & Special Characters', 'Mandatory', '1. Address Line 1 must have a value.\n2. Value must not exceed 500 characters.'],
    ['Address Line 2', 'Free Text', '1-500 Alpha Numeric & Special Characters', 'Optional', '1. Value must not exceed 500 characters.'],
    ['Address Line 3', 'Free Text', '1-500 Alpha Numeric & Special Characters', 'Optional', '1. Value must not exceed 500 characters.'],
    ['Company Email', 'Free Text', '1-100 Alpha Numeric & Special Characters', 'Optional', '1. If entered, Company Email must be a valid email format.\n2. Value must not exceed 100 characters.\n3. Duplicate Company Email across Leads is allowed (no uniqueness check).'],
    ['Company Contact No.', 'Free Text', '1-20 Numeric & Special Characters', 'Optional', '1. Value must not exceed 20 characters.\n2. Duplicate Company Contact No. across Leads is allowed (no uniqueness check).'],
    ['Company Website', 'Free Text', '1-100 Alpha Numeric & Special Characters', 'Optional', '1. If entered, Company Website must be a valid URL/website format.\n2. Value must not exceed 100 characters.'],
    ['Source', 'Dropdown (Real-time, Company Master)', 'Value from Dropdown', 'Optional', '1. Each cell of the column contains the Source options configured for the company. Source: FMS -> Master -> Others -> User Value Detail -> Lead Source.\n2. One of the options should be selected from the dropdown.\n3. If no option is selected or an incorrect value is entered, Source shall be empty for the Lead (no blockage).'],
    ['Lead Owner Email Id*', 'Free Text', '1-100 Alpha Numeric & Special Characters', 'Mandatory', '1. Lead Owner Email Id must have a value.\n2. Value must be a valid email format.\n3. Email ID must match the email of an existing user. Source: FMS -> Admin -> Company Details -> Manage My Users.\n4. Value must not exceed 100 characters.'],
    ['Lead Stage*', 'Dropdown (Real-time, Lead Stage Master)', 'Value from Dropdown', 'Mandatory', '1. Lead Stage must have a value.\n2. Each cell of the column contains the Lead Stage options configured under Sales -> Settings -> Lead Stages.\n3. Value must match one of the configured Lead Stage options exactly.'],
    ['Industry', 'Dropdown (Real-time, Company Master)', 'Value from Dropdown', 'Optional', '1. Each cell of the column contains the Industry options configured for the company. Source: FMS -> Master -> Others -> User Value Detail -> Industry.\n2. If no option is selected or an incorrect value is entered, Industry shall be empty for the Lead (no blockage).'],
    ['Lead Temperature', 'Dropdown (Real-time, Company Master)', 'Value from Dropdown', 'Optional', '1. Each cell of the column contains the Lead Temperature options configured for the company. Source: FMS -> Master -> Others -> User Value Detail -> Opportunity Rating.\n2. If left blank or an incorrect value is entered, Lead Temperature defaults to \'Hot\'.'],
    ['Service Interest', 'Dropdown (Real-time, Company Master)', 'Value from Dropdown', 'Optional', '1. Each cell of the column contains the Service Interest options configured for the company. Source: FMS -> Master -> Others -> User Value Detail -> Service Interest.\n2. If no option is selected or an incorrect value is entered, Service Interest shall be empty for the Lead (no blockage).'],
    ['Networks', 'Dropdown (Real-time, Company Master)', 'Value from Dropdown', 'Optional', '1. Each cell of the column contains the Networks options configured for the company. Source: FMS -> Master -> Others -> User Value Detail -> Networks.\n2. If no option is selected or an incorrect value is entered, Networks shall be empty for the Lead (no blockage).'],
    ['Notes', 'Free Text', '1-1000 Alpha Numeric & Special Characters', 'Optional', '1. Value must not exceed 1000 characters.'],
    ['Primary Contact', 'Boolean Dropdown', 'TRUE / FALSE', 'Optional', '1. Only TRUE, FALSE, or blank is accepted.\n2. If TRUE: First Name, Last Name and Email become Mandatory for that row. If any of the three is left blank, a mandatory-field validation is triggered for the missing field(s).\n3. If FALSE or left blank, any values entered in First Name, Last Name, Email, Phone No. (Primary) and Designation are ignored and not saved against the Lead - these fields are only processed when Primary Contact = TRUE.'],
    ['First Name', 'Free Text', '1-50 Alpha Numeric & Special Characters', 'Conditionally Mandatory (Mandatory when Primary Contact = TRUE)', '1. Mandatory when Primary Contact = TRUE; otherwise Optional/ignored.\n2. Value must not exceed 50 characters.'],
    ['Last Name', 'Free Text', '1-50 Alpha Numeric & Special Characters', 'Conditionally Mandatory (Mandatory when Primary Contact = TRUE)', '1. Mandatory when Primary Contact = TRUE; otherwise Optional/ignored.\n2. Value must not exceed 50 characters.'],
    ['Email', 'Free Text', '1-100 Alpha Numeric & Special Characters', 'Conditionally Mandatory (Mandatory when Primary Contact = TRUE)', '1. Mandatory when Primary Contact = TRUE; otherwise Optional/ignored.\n2. If entered, Email must be a valid email format.\n3. Value must not exceed 100 characters.'],
    ['Phone No. (Primary)', 'Free Text', '1-20 Numeric & Special Characters', 'Optional', '1. Only processed when Primary Contact = TRUE (otherwise ignored - see Primary Contact validations).\n2. Phone Number must contain 20 or fewer characters.\n3. Must contain digits and permitted special characters only (+, spaces, hyphens, parentheses) - no alphabets.'],
    ['Designation', 'Dropdown (Real-time, Company Master)', 'Value from Dropdown', 'Optional', '1. Only processed when Primary Contact = TRUE (otherwise ignored).\n2. Each cell of the column contains the Designation options configured for the company. Source: FMS -> Master -> Others -> User Value Detail -> Designation.\n3. If no option is selected or an incorrect value is entered, Designation shall be empty for the Lead (no blockage).']
  ];

  const ws2 = XLSX.utils.aoa_to_sheet([instructionHeaders, ...instructionRows]);
  ws2['!cols'] = [
    { wch: 22 },
    { wch: 25 },
    { wch: 35 },
    { wch: 28 },
    { wch: 70 }
  ];

  XLSX.utils.book_append_sheet(wb, ws1, 'Import Bulk Lead');
  XLSX.utils.book_append_sheet(wb, ws2, 'Instructions');

  XLSX.writeFile(wb, 'Lead_Bulk_Import_Template.xlsx');
  showToast('Template Downloaded', 'Lead_Bulk_Import_Template.xlsx with 2 sheets generated.', 'success');
}

// Quick Sample Loaders for Testing Scenarios directly in the browser
window.loadSampleScenario = function(type) {
  if (!importPermissionEnabled) {
    showToast(
      'Access Restricted',
      'You do not have permission to perform this action. Please contact your Admin or service@freightoscope.com for assistance.',
      'error'
    );
    return;
  }
  openDrawer();

  let sampleData = [];
  let fileName = 'Sample_Leads.xlsx';

  if (type === 'valid') {
    fileName = 'Valid_3_Leads.xlsx';
    sampleData = [
      DEFINED_HEADERS,
      [
        'Nexus Global Freight Ltd', 'NGF', 'Nexus Global Freight Limited', 'Customer', 'US', 'California', 'Los Angeles', '90001', '500 Harbor Dr', '', '', 'info@nexusfreight.com', '+1 310-555-1234', 'www.nexusfreight.com', 'Website', 'sakshi.barnwal@freightoscope.com', 'New', 'Freight Forwarding', 'Hot', 'Air Freight', 'IATA', 'Key account prospect', 'TRUE', 'Sharon', 'Roy', 'sharon.roy@nexusfreight.com', '+1 310-555-9876', 'Logistics Director'
      ],
      [
        'Blue Ocean Shipping LLC', 'BOS', 'Blue Ocean Shipping LLC', 'Customer', 'AE', 'Dubai', 'Dubai', '00000', 'Office 402, Business Bay Tower', '', '', 'ops@blueocean.ae', '+971 4 555 1234', 'www.blueocean.ae', 'Referral', 'rafia.khan@freightoscope.com', 'Qualified', 'Shipping', 'Hot', 'Ocean Freight (FCL)', 'FIATA', 'Ready for quote', 'FALSE', '', '', '', '', ''
      ],
      [
        'Apex World Transports Pvt', 'AWT', 'Apex World Transports Private Limited', 'Agent', 'IN', 'Maharashtra', 'Mumbai', '400001', '12 Nariman Point, Express Towers', '', '', 'contact@apexworld.in', '+91 22 5555 1234', 'www.apexworld.in', 'Sales Call', 'john.doe@freightoscope.com', 'New', 'Logistics', 'Warm', 'Customs Clearance', 'CII', 'Expansion in India ports', 'TRUE', 'Rajesh', 'Kumar', 'rajesh.k@apexworld.in', '+91 98200 12345', 'Managing Partner'
      ]
    ];
  } else if (type === 'errors') {
    fileName = 'Sample_With_Field_Errors.xlsx';
    sampleData = [
      DEFINED_HEADERS,
      [
        '', // Missing Company Name
        'NGF', 'Nexus Global', 'Customer',
        'USA', // Invalid 2-letter Country Code
        'CA', 'Los Angeles', '90001',
        '', // Missing Address Line 1
        '', '',
        'not-an-email', // Invalid Email format
        '',
        'http://nexus.com', '',
        'unknown.user@freightoscope.com', // Owner Not Found
        '', // Missing Lead Stage
        '', '', '', '', '',
        'TRUE',
        '', // Missing First Name (Primary Contact = TRUE)
        'Roy',
        'sharon-email', // Invalid Contact Email
        '+1 310-CALL-NOW', // Letters in phone number
        'Director'
      ],
      [
        'Global Express Cargo Corporation International Overseas Worldwide Global Line Limited Company Extra Name Exceeding Limit Very Long Name Over Hundred Characters Completely', // > 100 chars
        'GEC', 'Global Express', 'Customer',
        'ZZ', // Country Not Found
        'NY',
        'New York City Central Metropolitan Regional Area Exceeding The One Hundred Character Limit Length Test Checking Rule', // City > 100 chars
        '10001',
        '100 Broadway St', '', '',
        'info@globalexpress.com', '', 'www.globalexpress.com', '',
        'invalid-email-format', // Invalid Owner Email format
        'New', '', '', '', '', '',
        'MAYBE', // Invalid Primary Contact boolean value
        '', '', '', '', ''
      ]
    ];
  } else if (type === 'duplicates') {
    fileName = 'Sample_With_Duplicate_Leads.xlsx';
    sampleData = [
      DEFINED_HEADERS,
      [
        'Evergreen Freight Forwarders', 'EFF', 'Evergreen Freight Forwarders', 'Customer', 'IN', 'Tamilnadu', 'Chennai', '600001', '14 Rajaji Salai', '', '', 'info@evergreen.in', '', 'www.evergreen.in', '', 'sakshi.barnwal@freightoscope.com', 'New', '', 'Hot', '', '', '', 'FALSE', '', '', '', '', ''
      ],
      [
        'Pacific Trans Logistics', 'PTL', 'Pacific Trans', 'Customer', 'US', 'WA', 'Seattle', '98101', '400 Pine St', '', '', 'info@pacific.com', '', 'www.pacific.com', '', 'rafia.khan@freightoscope.com', 'Qualified', '', 'Hot', '', '', '', 'FALSE', '', '', '', '', ''
      ],
      [
        'Evergreen Freight Forwarders ', 'EFF', 'Evergreen Freight Forwarders', 'Customer', 'IN', 'Tamilnadu', ' Chennai ', '600001', '14 Rajaji Salai', '', '', 'contact@evergreen.in', '', 'www.evergreen.in', '', 'sakshi.barnwal@freightoscope.com', 'New', '', 'Hot', '', '', '', 'FALSE', '', '', '', '', ''
      ]
    ];
  } else if (type === 'empty') {
    fileName = 'Empty_Lead_Template.xlsx';
    sampleData = [
      DEFINED_HEADERS
      // No rows
    ];
  } else if (type === 'wrong_headers') {
    fileName = 'Corrupted_Headers.xlsx';
    sampleData = [
      ['Customer Name', 'City Location', 'Phone', 'Email Address'],
      ['Acme Corp', 'New York', '1234567', 'test@acme.com']
    ];
  }

  // Create virtual file and trigger parsing
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(sampleData);
  XLSX.utils.book_append_sheet(wb, ws, 'Import Bulk Lead');

  // UI simulate file
  uploadedFileName.textContent = fileName;
  uploadedFileSize.textContent = '14.8 KB';
  uploadedFileBar.style.display = 'flex';
  btnDrawerReset.style.display = 'inline-block';
  hideBanners();
  hideErrors();
  hideSuccessBox();

  processWorkbook(wb);
};

// Toast Notifications System
function showToast(title, message, type = 'info', duration = 4500) {
  const container = document.getElementById('toastContainer');
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;

  let icon = 'ℹ️';
  if (type === 'success') icon = '✅';
  if (type === 'error') icon = '❌';
  if (type === 'warning') icon = '⚠️';

  toast.innerHTML = `
    <div class="toast-icon">${icon}</div>
    <div class="toast-content">
      <div class="toast-title">${escapeHtml(title)}</div>
      <div class="toast-desc">${escapeHtml(message)}</div>
    </div>
    <button class="toast-close">&times;</button>
  `;

  toast.querySelector('.toast-close').addEventListener('click', () => {
    toast.remove();
  });

  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    toast.style.transition = 'all 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
