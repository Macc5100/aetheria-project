const REGISTRATION_SHEET_NAME = 'Registrations';
const SETTINGS_SHEET_NAME = 'Settings';
const DEFAULT_PIXEL_ID = '1443191528011039';
const ADMIN_TOKEN_TTL_SECONDS = 21600;

function doGet() {
  try {
    const settings = readSettings_();
    return jsonOutput_({
      status: 'success',
      announcement: settings.announcement,
      cbtStatus: settings.cbtStatus,
      pixelId: settings.pixelId,
      pixelStatus: settings.pixelStatus
    });
  } catch (error) {
    console.error(error);
    return jsonOutput_({ status: 'error', message: 'Unable to load public settings' });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('Request body is required');

    const data = JSON.parse(e.postData.contents);
    if (!data || typeof data !== 'object' || Array.isArray(data)) throw new Error('Invalid request body');

    switch (data.action) {
      case 'adminLogin':
        return adminLogin_(data);
      case 'getRegistrations':
        requireAdmin_(data.token);
        return getRegistrations_();
      case 'updateSettings':
        requireAdmin_(data.token);
        return updateSettings_(data);
      case 'adminLogout':
        logoutAdmin_(data.token);
        return jsonOutput_({ status: 'success' });
      case undefined:
        return registerUser_(data);
      default:
        throw new Error('Unknown action');
    }
  } catch (error) {
    console.error(error);
    return jsonOutput_({ status: 'error', message: error.message || 'Request failed' });
  }
}

function adminLogin_(data) {
  const properties = PropertiesService.getScriptProperties();
  const configuredPassword = properties.getProperty('ADMIN_PASSWORD');
  if (!configuredPassword) {
    return jsonOutput_({ status: 'error', message: 'Admin password is not configured in Script Properties' });
  }
  if (typeof data.password !== 'string' || data.password !== configuredPassword) {
    return jsonOutput_({ status: 'error', message: 'Invalid admin password' });
  }

  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('admin_' + token, 'authorized', ADMIN_TOKEN_TTL_SECONDS);
  return jsonOutput_({ status: 'success', token: token, expiresIn: ADMIN_TOKEN_TTL_SECONDS });
}

function requireAdmin_(token) {
  if (typeof token !== 'string' || !CacheService.getScriptCache().get('admin_' + token)) {
    throw new Error('Unauthorized');
  }
}

function logoutAdmin_(token) {
  if (typeof token === 'string') CacheService.getScriptCache().remove('admin_' + token);
}

function registerUser_(data) {
  const settings = readSettings_();
  if (settings.cbtStatus !== 'open') throw new Error('Registration is closed');

  const email = String(data.email || '').trim().toLowerCase();
  const phone = String(data.phone || '').trim();
  if (email.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error('A valid email address is required');
  }
  if (!/^\d{9,10}$/.test(phone)) throw new Error('A valid phone number is required');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // Recheck after acquiring the lock so closing registration cannot race with a submission.
    if (readSettings_().cbtStatus !== 'open') throw new Error('Registration is closed');

    const sheet = getRegistrationSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const existing = sheet.getRange(2, 2, lastRow - 1, 2).getDisplayValues();
      const duplicate = existing.some(row =>
        String(row[0] || '').trim().toLowerCase() === email || String(row[1] || '').trim() === phone
      );
      if (duplicate) throw new Error('Email or phone is already registered');
    }

    sheet.getRange(sheet.getLastRow() + 1, 2, 1, 2).setNumberFormat('@');
    sheet.appendRow([new Date(), email, phone]);
  } finally {
    lock.releaseLock();
  }

  return jsonOutput_({ status: 'success', message: 'Registration saved' });
}

function getRegistrations_() {
  const sheet = getRegistrationSheet_();
  const values = sheet.getDataRange().getValues();
  const registrations = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i][1] || values[i][2]) {
      registrations.push({
        timestamp: values[i][0] instanceof Date ? values[i][0].toISOString() : String(values[i][0] || ''),
        email: String(values[i][1] || ''),
        phone: String(values[i][2] || '')
      });
    }
  }

  const settings = readSettings_();
  return jsonOutput_({
    status: 'success',
    registrations: registrations,
    announcement: settings.announcement,
    cbtStatus: settings.cbtStatus,
    pixelId: settings.pixelId,
    pixelStatus: settings.pixelStatus
  });
}

function updateSettings_(data) {
  const announcement = String(data.announcement || '').trim();
  const cbtStatus = data.cbtStatus;
  const pixelId = String(data.pixelId || '').trim();
  const pixelStatus = data.pixelStatus;

  if (announcement.length > 500) throw new Error('Announcement must be 500 characters or fewer');
  if (cbtStatus !== 'open' && cbtStatus !== 'closed') throw new Error('Invalid CBT status');
  if (pixelStatus !== 'active' && pixelStatus !== 'inactive') throw new Error('Invalid Pixel status');
  if (pixelId && !/^\d{5,20}$/.test(pixelId)) throw new Error('Pixel ID must contain 5 to 20 digits');
  if (pixelStatus === 'active' && !pixelId) throw new Error('Pixel ID is required while Pixel is active');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getOrCreateSheet_(SETTINGS_SHEET_NAME);
    sheet.getRange('A1:D1').setValues([[announcement, cbtStatus, pixelId, pixelStatus]]);
  } finally {
    lock.releaseLock();
  }
  return jsonOutput_({ status: 'success' });
}

function readSettings_() {
  const sheet = getSpreadsheet_().getSheetByName(SETTINGS_SHEET_NAME);
  const hasSavedSettings = sheet && sheet.getLastRow() > 0;
  if (!hasSavedSettings) {
    return {
      announcement: '',
      cbtStatus: 'open',
      pixelId: DEFAULT_PIXEL_ID,
      pixelStatus: 'inactive'
    };
  }

  const values = sheet.getRange('A1:D1').getValues()[0];
  return {
    announcement: String(values[0] || ''),
    cbtStatus: values[1] === 'closed' ? 'closed' : 'open',
    pixelId: String(values[2] || DEFAULT_PIXEL_ID),
    pixelStatus: values[3] === 'active' && values[2] ? 'active' : 'inactive'
  };
}

function getRegistrationSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(REGISTRATION_SHEET_NAME);
  if (!sheet) {
    const legacySheet = ss.getSheets().find(candidate => {
      if (candidate.getName() === SETTINGS_SHEET_NAME || candidate.getLastRow() === 0) return false;
      const columnCount = Math.min(3, candidate.getLastColumn());
      if (columnCount < 3) return false;
      const headers = candidate.getRange(1, 1, 1, 3).getDisplayValues()[0]
        .map(value => String(value).trim().toLowerCase());
      return headers[0] === 'timestamp' && headers[1] === 'email' &&
        (headers[2] === 'phone' || headers[2] === 'phone number');
    });
    sheet = legacySheet || ss.insertSheet(REGISTRATION_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) sheet.appendRow(['Timestamp', 'Email', 'Phone Number']);
  return sheet;
}

function getOrCreateSheet_(name) {
  const ss = getSpreadsheet_();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function getSpreadsheet_() {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
  if (spreadsheetId) return SpreadsheetApp.openById(spreadsheetId);

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error('Set SPREADSHEET_ID in Script Properties');
  return active;
}

function jsonOutput_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
