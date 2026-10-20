/** Kalibrasi KM: entry point, authentication, database, unit workflow and triggers. */


// ============================================================================
// 01. WEB ENTRY & BOOTSTRAP
// ============================================================================
/** Web entry point, template partials, bootstrap, and database setup orchestration. */
function include_(fileName) {
  const allowedPartials = ['Styles'];
  if (!allowedPartials.includes(fileName)) {
    throw new Error('Partial HTML tidak diizinkan: ' + fileName);
  }
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

function doGet() {
  // Pada deployment Execute as Me + Anyone, halaman selalu dibuka sebagai login gate.
  // Seluruh data/endpoint tetap terkunci oleh server-side session + RBAC.
  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle('Monitoring Kalibrasi KM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  const actor = getCurrentUserContext_({ allowUnauthorized: true, source: 'SHEET' });
  if (actor.authorized && actor.permissions.includes('ADMIN_SETUP')) {
    SpreadsheetApp.getUi().createMenu('Kalibrasi KM')
      .addItem('Setup / Migrasi Monitoring', 'setupDatabaseKalibrasiKMFromSheet')
      .addItem('Audit Database (Read-only)', 'auditDatabaseKalibrasiKMFromSheet')
      .addItem('Audit Relasi Foto (Read-only)', 'auditCalibrationPhotoIntegrityFromSheet')
      .addItem('Audit Drift Status 3L (Read-only)', 'auditExternalStatus3LDriftFromSheet')
      .addItem('Repair Workflow Status 3L', 'repairStatus3LWorkflowIntegrity')
      .addItem('Sinkronkan Status 3L Eksternal', 'repairExternalStatus3LChangesFromSheet')
      .addItem('Rebuild Analytics Ledger', 'rebuildAnalyticsCalibrationLedgerFromSheet')
      .addItem('Rapikan Struktur + Backup', 'cleanDatabaseStructureFromSheet').addToUi();
  }
}

function setupDatabaseKalibrasiKM(sessionToken) {
  const actor = requirePermission_('ADMIN_SETUP', '', { token: sessionToken });
  return setupDatabaseKalibrasiKMCore_(actor);
}

function repairDatabaseKalibrasiKM(sessionToken) {
  const actor = requirePermission_('ADMIN_SETUP', '', { token: sessionToken });
  const result = setupDatabaseKalibrasiKMCore_(actor);
  return {
    success: true,
    migration: result.migration || {},
    masterSetup: result.masterSetup || {},
    inspection: inspectDatabaseKalibrasiKm_()
  };
}

function requireSheetMaintenanceAdmin_() {
  const activeEmail = getSheetUserEmail_();
  const ownerEmail = getSpreadsheetOwnerEmail_();
  if (!activeEmail) {
    throw new Error('AUTH_REQUIRED: email akun Apps Script tidak dapat dibaca. Jalankan fungsi ini dari editor Apps Script menggunakan akun owner Spreadsheet.');
  }
  if (ownerEmail && activeEmail !== ownerEmail) {
    throw new Error('ACCESS_DENIED: maintenance database hanya boleh dijalankan oleh owner Spreadsheet.');
  }
  const actor = requirePermission_('ADMIN_SETUP', '', { source: 'SHEET' });
  KM_REQUEST_ACTOR_ = actor;
  return actor;
}

/**
 * Fungsi maintenance PUBLIK untuk dijalankan langsung dari dropdown Apps Script.
 * Tidak memerlukan sessionToken Web App.
 */
function setupDatabaseKalibrasiKMFromSheet() {
  const actor = requireSheetMaintenanceAdmin_();
  return setupDatabaseKalibrasiKMCore_(actor);
}

function cleanDatabaseStructureFromSheet() {
  return cleanDatabaseStructure_(requireSheetMaintenanceAdmin_());
}

/** Alias internal / kompatibilitas menu versi lama. */
function setupDatabaseKalibrasiKMFromSheet_() {
  return setupDatabaseKalibrasiKMFromSheet();
}

/**
 * Repair + audit database dari Apps Script editor tanpa sessionToken Web App.
 */
function repairDatabaseKalibrasiKMFromSheet() {
  const actor = requireSheetMaintenanceAdmin_();
  const result = setupDatabaseKalibrasiKMCore_(actor);
  return {
    success: true,
    migration: result.migration || {},
    masterSetup: result.masterSetup || {},
    inspection: inspectDatabaseKalibrasiKm_()
  };
}

/**
 * Audit schema database dari Apps Script editor tanpa sessionToken Web App.
 */
function auditDatabaseKalibrasiKMFromSheet() {
  requireSheetMaintenanceAdmin_();
  return inspectDatabaseKalibrasiKm_();
}

function setupDatabaseKalibrasiKMCore_(actor) {
  KM_REQUEST_ACTOR_ = actor;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const preMigrationBackup = databaseNeedsMigrationBackup_(ss)
    ? createAdministrationBackup_(actor, 'BEFORE_SCHEMA_V22_MIGRATION') : null;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let migration = null;
  try {
    migration = migrateMonitoringDatabase_(ss);
    setupRbacDatabase_(ss);
    setupConfigSheet_(getOrCreateSheet_(ss, KM_APP.SHEETS.CONFIG));
    setupLogSheet_(getOrCreateSheet_(ss, KM_APP.SHEETS.LOG));
    SpreadsheetApp.flush();
    invalidateAppCaches_();
    invalidateSecurityCaches_();
  } finally { lock.releaseLock(); }
  const masterSetup = setupMasterUnitKmDatabase_();
  ensureAnalyticsCalibrationSheet_(ss);
  let analyticsLedger = null;
  runNonCritical_('sinkronisasi clean analytics ledger', function () {
    const audit = auditAnalyticsCalibrationLedger_();
    analyticsLedger = audit.ready
      ? { success: true, rebuilt: false, audit: audit }
      : Object.assign({ rebuilt: true }, rebuildAnalyticsCalibrationLedgerCore_(actor), { auditBefore: audit });
  });
  const headerOrder = runNonCritical_('standardisasi urutan header enterprise', function () {
    return standardizeEnterpriseHeaderOrder_(ss);
  });
  // Proteksi Google Sheet tidak dipakai sebagai boundary akses. Bersihkan
  // proteksi lama buatan aplikasi; RBAC tetap diverifikasi di backend.
  runNonCritical_('hapus proteksi otomatis lama', removeAutomaticSheetProtections_);
  KM_REQUEST_MIGRATION_DIRECTORY_ = null;
  KM_REQUEST_STATUS3L_OPTIONS_ = null;
  const compacted = runNonCritical_('kompak grid database', compactDatabaseGrid_);
  runNonCritical_('normalisasi format tanggal seluruh database', function () {
    formatAllSheetDateColumns_(SpreadsheetApp.getActiveSpreadsheet());
    formatDatabaseTables_(SpreadsheetApp.getActiveSpreadsheet());
  });
  return {
    success: true,
    schemaVersion: KM_DATABASE_SCHEMA_VERSION,
    preMigrationBackup: preMigrationBackup,
    migration: migration,
    masterSetup: masterSetup,
    analyticsLedger: analyticsLedger,
    headerOrder: headerOrder || [],
    compacted: compacted,
    inventory: auditDatabaseSheetInventory_(ss)
  };
}

function setupConfigSheet_(sheet) {
  const params = getExistingParameters_(sheet);
  const status3LOptions = (KM_REQUEST_STATUS3L_OPTIONS_ || []).slice().concat(
    getConfigColumnValuesByHeader_(sheet, 'STATUS PROSES 3L')
  );
  const pageSize = [15, 30, 50].includes(Number(params.PAGE_SIZE)) ? Number(params.PAGE_SIZE) : 15;
  const appName = params.APP_NAME || 'Monitoring Kalibrasi KM';
  const approvalEnabled = ['TRUE','1','YES','ON','ENABLED'].includes(normalizeLookupText_(params[MASTER_UNIT_MANAGEMENT.APPROVAL_PARAMETER] || 'FALSE'));
  const parameterRows = buildManagedConfigParameterRows_(params, {
    SCHEMA_VERSION: String(KM_DATABASE_SCHEMA_VERSION),
    PAGE_SIZE: pageSize,
    APP_NAME: appName,
    [MASTER_UNIT_MANAGEMENT.APPROVAL_PARAMETER]: approvalEnabled ? 'TRUE' : 'FALSE'
  });
  const rowCount = Math.max(6, sheet.getLastRow(), status3LOptions.length + 1,
    parameterRows.length + 1, (KM_REQUEST_MIGRATION_DIRECTORY_ || []).length + 1);
  if (sheet.getMaxRows() < rowCount) sheet.insertRowsAfter(sheet.getMaxRows(), rowCount - sheet.getMaxRows());
  if (sheet.getMaxColumns() < 6) sheet.insertColumnsAfter(sheet.getMaxColumns(), 6 - sheet.getMaxColumns());
  // Schema v22: A:C directory, D Status 3L, E:F parameter aplikasi.
  sheet.getRange(1, 1, sheet.getMaxRows(), Math.min(sheet.getMaxColumns(), 6)).clearContent().clearDataValidations();
  sheet.getRange('A1:F1').setValues([[
    'OPERATING POINT', 'NAMA OPC/WHC', 'NAMA TOH/WOH', 'STATUS PROSES 3L', 'PARAMETER', 'VALUE'
  ]]);
  const directory = KM_REQUEST_MIGRATION_DIRECTORY_ || [];
  if (directory.length) sheet.getRange(2, 1, directory.length, 3).setValues(directory);
  const statuses = uniqueClean_(status3LOptions.concat(DEFAULT_STATUS_3L_OPTIONS));
  sheet.getRange(2, 4, statuses.length, 1).setValues(statuses.map(value => [safeSheetText_(value)]));
  if (parameterRows.length) sheet.getRange(2, 5, parameterRows.length, 2).setValues(parameterRows);
  if (sheet.getMaxColumns() > 6) sheet.deleteColumns(7, sheet.getMaxColumns() - 6);
  formatConfigSheet_(sheet);
}

function getConfigColumnValuesByHeader_(sheet, headerName) {
  if (!sheet || !sheet.getLastColumn() || !sheet.getLastRow()) return [];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const index = headers.findIndex(header =>
    normalizeHeaderKey_(header) === normalizeHeaderKey_(headerName)
  );
  if (index < 0 || sheet.getLastRow() < 2) return [];
  return uniqueClean_(sheet.getRange(2, index + 1, sheet.getLastRow() - 1, 1)
    .getDisplayValues().map(row => row[0]));
}

function getExistingParameters_(sheet) {
  const result = {};
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 2) return result;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  let parameterColumn = headers.findIndex(header => normalizeHeaderKey_(header) === 'PARAMETER') + 1;
  let valueColumn = headers.findIndex(header => normalizeHeaderKey_(header) === 'VALUE') + 1;

  // Fallback hanya untuk database lama sebelum migrasi.
  if (!parameterColumn || !valueColumn) {
    if (sheet.getLastColumn() >= 14) {
      parameterColumn = 13;
      valueColumn = 14;
    } else {
      return result;
    }
  }

  const rowCount = Math.max(sheet.getLastRow() - 1, 1);
  const keys = sheet.getRange(2, parameterColumn, rowCount, 1).getDisplayValues();
  const vals = sheet.getRange(2, valueColumn, rowCount, 1).getDisplayValues();
  keys.forEach((row, index) => {
    const key = String(row[0] || '').trim();
    const value = String(vals[index] && vals[index][0] || '').trim();
    if (key) result[key] = value;
  });
  return result;
}

/**
 * Menulis ulang parameter yang dikelola aplikasi tanpa menghapus parameter
 * tambahan milik integrasi/versi lain. Override selalu ditempatkan lebih dulu
 * agar hasil sheet mudah diaudit dan hanya memiliki satu baris per key.
 */
function buildManagedConfigParameterRows_(existing, overrides) {
  const source = existing || {};
  const managed = overrides || {};
  const rows = [];
  const written = Object.create(null);
  Object.keys(managed).forEach(key => {
    const normalizedKey = cleanText_(key);
    if (!normalizedKey) return;
    rows.push([safeSheetText_(normalizedKey), safeSheetText_(managed[key])]);
    written[normalizeLookupText_(normalizedKey)] = true;
  });
  Object.keys(source).sort().forEach(key => {
    const normalizedKey = cleanText_(key);
    if (!normalizedKey || written[normalizeLookupText_(normalizedKey)]) return;
    rows.push([safeSheetText_(normalizedKey), safeSheetText_(source[key])]);
    written[normalizeLookupText_(normalizedKey)] = true;
  });
  return rows;
}

function setupLogSheet_(sheet) {
  ensureHeaders_(sheet, KM_APP.LOG_HEADERS);
}

function getAppBootstrapData(sessionToken) {
  const actor = requirePermission_('ACCESS_APP', '', { token: sessionToken });
  return buildAppBootstrapData_(actor);
}


function buildAppBootstrapData_(actor) {
  if (!actor || !actor.authorized || !actor.permissions.includes('ACCESS_APP')) {
    throw new Error('ACCESS_DENIED: akun tidak memiliki izin membuka aplikasi.');
  }
  KM_REQUEST_ACTOR_ = actor;
  const configuredPageSize = Number(getConfigParameter_('PAGE_SIZE', 15));

  return {
    masterOptions: getMasterDropdownData_(),
    config: getConfigLists_(),
    currentUser: actor.email,
    userContext: publicActorContext_(actor),
    rbacRoles: readRoleDirectory_().filter(role => role.active).map(role => role.role),
    serverToday: dateOnlyKey_(new Date()),
    submissionSchedule: getCalibrationSubmissionSchedule_(),
    appName: getConfigParameter_('APP_NAME', 'Monitoring Kalibrasi KM'),
    pageSize: [15, 30, 50].includes(configuredPageSize)
      ? configuredPageSize
      : 15,
  };
}

function getMasterDropdownData(sessionToken) {
  requirePermission_('ACCESS_APP', '', { token: sessionToken });
  return getMasterDropdownData_();
}

function getMasterDropdownData_() {
  const actor = requirePermission_('ACCESS_APP');
  const empty = { operatingPoint: [], opc: [], toh: [], status3L: [] };
  const scopeResult = result => Object.assign({}, result, {
    operatingPoint: actor.globalScope
      ? (result.operatingPoint || []).slice()
      : (result.operatingPoint || []).filter(value => canAccessOperatingPoint_(actor, value))
  });
  const cached = getCachedJson_('master-dropdowns');
  if (cached) return scopeResult(cached);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.CONFIG);
  if (!sheet || sheet.getLastRow() < 2) return scopeResult(empty);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const indexOf = headerName => headers.findIndex(header => normalizeHeaderKey_(header) === normalizeHeaderKey_(headerName));
  const opIndex = indexOf('OPERATING POINT');
  const opcIndex = indexOf('NAMA OPC/WHC');
  const tohIndex = indexOf('NAMA TOH/WOH');
  const statusIndex = indexOf('STATUS PROSES 3L');
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
  const result = {
    operatingPoint: opIndex >= 0 ? uniqueClean_(values.map(row => row[opIndex])) : [],
    opc: opcIndex >= 0 ? uniqueClean_(values.map(row => row[opcIndex])) : [],
    toh: tohIndex >= 0 ? uniqueClean_(values.map(row => row[tohIndex])) : [],
    status3L: statusIndex >= 0 ? uniqueClean_(values.map(row => row[statusIndex])) : []
  };
  putCachedJson_('master-dropdowns', result, 120);
  return scopeResult(result);
}

function getConfigLists_() {
  return {
    unitStatuses: UNIT_MANUAL_OPERATIONAL_STATUSES.slice(),
    operationalStatuses: UNIT_OPERATIONAL_STATUSES.slice(),
    status3LOptions: getStatus3LOptions_()
  };
}

let KM_REQUEST_STATUS_OPTIONS_ = null;
let KM_REQUEST_CONFIG_PARAMETERS_ = null;

const KM_CALIBRATION_SUBMISSION_SCHEDULE = Object.freeze({
  ENABLED: 'CALIBRATION_SUBMISSION_SCHEDULE_ENABLED',
  START_AT: 'CALIBRATION_SUBMISSION_START_AT',
  END_AT: 'CALIBRATION_SUBMISSION_END_AT',
  DAILY_START: 'CALIBRATION_SUBMISSION_DAILY_START',
  DAILY_END: 'CALIBRATION_SUBMISSION_DAILY_END',
  ALLOWED_DAYS: 'CALIBRATION_SUBMISSION_ALLOWED_DAYS'
});

function normalizeScheduleDateTime_(value, label) {
  const text = String(value == null ? '' : value).trim();
  if (!text) return '';
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})$/);
  if (!match || !createValidDateTime_(Number(match[1]), Number(match[2]), Number(match[3]),
      Number(match[4]), Number(match[5]), 0)) {
    throw new Error((label || 'Tanggal/waktu') + ' harus memakai format YYYY-MM-DD HH:mm yang valid.');
  }
  return match[1] + '-' + match[2] + '-' + match[3] + 'T' + match[4] + ':' + match[5];
}

function normalizeScheduleTime_(value, label) {
  const text = String(value == null ? '' : value).trim();
  if (!text) return '';
  const match = text.match(/^([01]\d|2[0-3]):([0-5]\d)$/);
  if (!match) throw new Error((label || 'Jam') + ' harus memakai format HH:mm.');
  return match[1] + ':' + match[2];
}

function normalizeScheduleAllowedDays_(value) {
  const source = Array.isArray(value) ? value : String(value == null ? '' : value).split(',');
  const days = [];
  source.forEach(item => {
    const day = Number(String(item).trim());
    if (!Number.isInteger(day) || day < 1 || day > 7) {
      if (String(item).trim()) throw new Error('Hari pengajuan hanya boleh bernilai 1 (Senin) sampai 7 (Minggu).');
      return;
    }
    if (!days.includes(day)) days.push(day);
  });
  return days.sort((a, b) => a - b);
}

function normalizeCalibrationSubmissionSchedule_(input) {
  const data = input || {};
  const enabled = data.enabled === true || ['TRUE', '1', 'YES', 'ON', 'ENABLED']
    .includes(normalizeLookupText_(data.enabled));
  const startAt = normalizeScheduleDateTime_(data.startAt, 'Awal periode pengajuan');
  const endAt = normalizeScheduleDateTime_(data.endAt, 'Akhir periode pengajuan');
  const dailyStart = normalizeScheduleTime_(data.dailyStart, 'Jam mulai harian');
  const dailyEnd = normalizeScheduleTime_(data.dailyEnd, 'Jam selesai harian');
  const allowedDays = normalizeScheduleAllowedDays_(data.allowedDays);
  if (startAt && endAt && startAt >= endAt) {
    throw new Error('Akhir periode pengajuan harus setelah awal periode.');
  }
  if (Boolean(dailyStart) !== Boolean(dailyEnd)) {
    throw new Error('Jam mulai dan jam selesai harian harus diisi berpasangan.');
  }
  if (dailyStart && dailyStart >= dailyEnd) {
    throw new Error('Jam selesai harian harus setelah jam mulai pada hari yang sama.');
  }
  return {
    enabled: enabled,
    startAt: startAt,
    endAt: endAt,
    dailyStart: dailyStart,
    dailyEnd: dailyEnd,
    allowedDays: allowedDays
  };
}

function calibrationSubmissionScheduleFromParameters_(params) {
  const values = params || {};
  return normalizeCalibrationSubmissionSchedule_({
    enabled: values[KM_CALIBRATION_SUBMISSION_SCHEDULE.ENABLED],
    startAt: values[KM_CALIBRATION_SUBMISSION_SCHEDULE.START_AT],
    endAt: values[KM_CALIBRATION_SUBMISSION_SCHEDULE.END_AT],
    dailyStart: values[KM_CALIBRATION_SUBMISSION_SCHEDULE.DAILY_START],
    dailyEnd: values[KM_CALIBRATION_SUBMISSION_SCHEDULE.DAILY_END],
    allowedDays: values[KM_CALIBRATION_SUBMISSION_SCHEDULE.ALLOWED_DAYS]
  });
}

function calibrationSubmissionScheduleParameterOverrides_(schedule) {
  const normalized = normalizeCalibrationSubmissionSchedule_(schedule);
  const result = {};
  result[KM_CALIBRATION_SUBMISSION_SCHEDULE.ENABLED] = normalized.enabled ? 'TRUE' : 'FALSE';
  result[KM_CALIBRATION_SUBMISSION_SCHEDULE.START_AT] = normalized.startAt;
  result[KM_CALIBRATION_SUBMISSION_SCHEDULE.END_AT] = normalized.endAt;
  result[KM_CALIBRATION_SUBMISSION_SCHEDULE.DAILY_START] = normalized.dailyStart;
  result[KM_CALIBRATION_SUBMISSION_SCHEDULE.DAILY_END] = normalized.dailyEnd;
  result[KM_CALIBRATION_SUBMISSION_SCHEDULE.ALLOWED_DAYS] = normalized.allowedDays.join(',');
  return result;
}

function scheduleLocalParts_(date, timeZone) {
  const instant = date instanceof Date ? date : new Date();
  const zone = timeZone || Session.getScriptTimeZone() || 'Asia/Jakarta';
  const localDate = Utilities.formatDate(instant, zone, 'yyyy-MM-dd');
  const localTime = Utilities.formatDate(instant, zone, 'HH:mm');
  const dateParts = localDate.split('-').map(Number);
  const jsDay = new Date(Date.UTC(dateParts[0], dateParts[1] - 1, dateParts[2])).getUTCDay();
  return {
    date: localDate,
    time: localTime,
    dateTime: localDate + 'T' + localTime,
    isoDay: jsDay === 0 ? 7 : jsDay,
    timeZone: zone
  };
}

function scheduleDayLabels_(days) {
  const labels = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'];
  return (days || []).map(day => labels[Number(day) - 1]).filter(Boolean);
}

function scheduleWindowLabel_(schedule) {
  const parts = [];
  if (schedule.startAt || schedule.endAt) {
    parts.push('periode ' + (schedule.startAt ? schedule.startAt.replace('T', ' ') : 'tanpa batas awal') +
      ' s.d. ' + (schedule.endAt ? schedule.endAt.replace('T', ' ') : 'tanpa batas akhir'));
  }
  if (schedule.allowedDays && schedule.allowedDays.length) {
    parts.push(scheduleDayLabels_(schedule.allowedDays).join(', '));
  }
  if (schedule.dailyStart && schedule.dailyEnd) {
    parts.push('pukul ' + schedule.dailyStart + '-' + schedule.dailyEnd);
  }
  return parts.join(' | ') || 'tanpa batas tanggal/jam';
}

function evaluateCalibrationSubmissionSchedule_(schedule, now) {
  const normalized = normalizeCalibrationSubmissionSchedule_(schedule);
  const parts = scheduleLocalParts_(now || new Date(), Session.getScriptTimeZone() || 'Asia/Jakarta');
  let isOpen = true;
  let reasonCode = normalized.enabled ? 'OPEN' : 'DISABLED';
  if (normalized.enabled && normalized.startAt && parts.dateTime < normalized.startAt) {
    isOpen = false;
    reasonCode = 'BEFORE_PERIOD';
  } else if (normalized.enabled && normalized.endAt && parts.dateTime > normalized.endAt) {
    isOpen = false;
    reasonCode = 'AFTER_PERIOD';
  } else if (normalized.enabled && normalized.allowedDays.length && !normalized.allowedDays.includes(parts.isoDay)) {
    isOpen = false;
    reasonCode = 'DAY_NOT_ALLOWED';
  } else if (normalized.enabled && normalized.dailyStart &&
      (parts.time < normalized.dailyStart || parts.time > normalized.dailyEnd)) {
    isOpen = false;
    reasonCode = parts.time < normalized.dailyStart ? 'BEFORE_DAILY_WINDOW' : 'AFTER_DAILY_WINDOW';
  }
  const windowLabel = scheduleWindowLabel_(normalized);
  return Object.assign({}, normalized, {
    isOpen: isOpen,
    reasonCode: reasonCode,
    timeZone: parts.timeZone,
    serverNow: parts.dateTime,
    windowLabel: windowLabel,
    message: !normalized.enabled
      ? 'Pengajuan kalibrasi tersedia setiap saat.'
      : isOpen
        ? 'Periode pengajuan sedang dibuka (' + windowLabel + ').'
        : 'Pengajuan kalibrasi sedang ditutup. Jadwal aktif: ' + windowLabel + '.'
  });
}

function getCalibrationSubmissionSchedule_() {
  const params = getConfigParameters_();
  try {
    return evaluateCalibrationSubmissionSchedule_(calibrationSubmissionScheduleFromParameters_(params), new Date());
  } catch (error) {
    const enabled = ['TRUE', '1', 'YES', 'ON', 'ENABLED'].includes(normalizeLookupText_(
      params[KM_CALIBRATION_SUBMISSION_SCHEDULE.ENABLED]
    ));
    return {
      enabled: enabled,
      isOpen: !enabled,
      reasonCode: 'INVALID_CONFIGURATION',
      timeZone: Session.getScriptTimeZone() || 'Asia/Jakarta',
      serverNow: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Jakarta', "yyyy-MM-dd'T'HH:mm"),
      startAt: String(params[KM_CALIBRATION_SUBMISSION_SCHEDULE.START_AT] || ''),
      endAt: String(params[KM_CALIBRATION_SUBMISSION_SCHEDULE.END_AT] || ''),
      dailyStart: String(params[KM_CALIBRATION_SUBMISSION_SCHEDULE.DAILY_START] || ''),
      dailyEnd: String(params[KM_CALIBRATION_SUBMISSION_SCHEDULE.DAILY_END] || ''),
      allowedDays: [],
      windowLabel: 'konfigurasi tidak valid',
      configurationError: error.message,
      message: enabled
        ? 'Pengajuan ditutup karena konfigurasi jadwal tidak valid. Hubungi administrator.'
        : 'Jadwal nonaktif; pengajuan tersedia setiap saat.'
    };
  }
}

function assertCalibrationSubmissionOpen_() {
  const schedule = getCalibrationSubmissionSchedule_();
  if (!schedule.isOpen) {
    throw new Error('CALIBRATION_SUBMISSION_CLOSED: ' + schedule.message);
  }
  return schedule;
}

function getStatus3LOptions_() {
  if (KM_REQUEST_STATUS_OPTIONS_) return KM_REQUEST_STATUS_OPTIONS_.slice();
  const cached = getCachedJson_('status-3l-options');
  if (cached && cached.length) {
    KM_REQUEST_STATUS_OPTIONS_ = cached;
    return cached.slice();
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.CONFIG);
  let options = [];
  if (sheet && sheet.getLastColumn() >= 1 && sheet.getLastRow() >= 2) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const columnIndex = headers.findIndex(header => normalizeHeaderKey_(header) === normalizeHeaderKey_('STATUS PROSES 3L'));
    if (columnIndex >= 0) {
      options = uniqueClean_(sheet.getRange(2, columnIndex + 1, sheet.getLastRow() - 1, 1)
        .getDisplayValues().map(row => row[0]));
    }
  }
  // Semua tahap canonical selalu tersedia pada dropdown dan summary, sambil
  // tetap mempertahankan status tambahan yang memang dikonfigurasi pengguna.
  options = uniqueClean_(options.concat(DEFAULT_STATUS_3L_OPTIONS));
  putCachedJson_('status-3l-options', options, 120);
  KM_REQUEST_STATUS_OPTIONS_ = options;
  return options.slice();
}

function resolveStatus3L_(input) {
  const options = getStatus3LOptions_();
  if (!options.length) throw new Error('MASTER_CONFIG belum memiliki opsi pada kolom STATUS PROSES 3L.');
  return resolveMasterValue_(options, input, 'Status 3L');
}

function resolveMasterValue_(list, input, label) {
  const value = cleanText_(input);

  if (!value) {
    throw new Error(label + ' wajib diisi.');
  }

  const exact = list.find(item => item === value);

  if (exact) {
    return exact;
  }

  const key = value.toUpperCase();
  const match = list.find(item => item.toUpperCase() === key);

  if (match) {
    return match;
  }

  throw new Error(
    label + ' tidak ditemukan pada MASTER_CONFIG: "' + value + '"'
  );
}

function isMasterUnitStatusValidationReady_(sheet, statusUnitIndex) {
  if (!sheet || statusUnitIndex < 0 || sheet.getMaxRows() < 2) return false;
  const validation = sheet.getRange(2, statusUnitIndex + 1).getDataValidation();
  if (!validation || validation.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    return false;
  }
  const criteria = validation.getCriteriaValues();
  const allowed = Array.isArray(criteria[0]) ? criteria[0].map(normalizeLookupText_) : [];
  // Ready to Depart harus menjadi anggota dropdown yang sah agar cell hasil trigger
  // tidak ditandai Invalid oleh Google Sheets. Walau terlihat di dropdown, perubahan
  // manual ke status ini tetap dijaga oleh workflow: hanya sah bila Status 3L = DONE.
  return !validation.getAllowInvalid() &&
    UNIT_OPERATIONAL_STATUSES.every(status => allowed.includes(normalizeLookupText_(status)));
}

function getMasterUnitData(forceRefresh, sessionToken) {
  const actor = requirePermission_('VIEW_MASTER_UNIT', '', { token: sessionToken });
  let revision = getCachedJson_('analytics-revision');
  if (!revision) { revision = Utilities.getUuid(); putCachedJson_('analytics-revision', revision, 600); }
  const responseCacheKey = 'master-scoped:' + sha256Hex_(JSON.stringify([revision, actor.email, actor.roles, actor.permissions, actor.globalScope, actor.operatingPointKeys || actor.operatingPoints]));
  if (!forceRefresh) {
    const cached = getCachedJson_(responseCacheKey);
    if (cached) return Object.assign({}, cached, {
      submissionSchedule: getCalibrationSubmissionSchedule_()
    });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);

  if (!sheet) {
    return scopeMasterUnitResponse_({
      rows: [],
      meta: {
        sheetFound: false,
        headers: [],
        totalRows: 0,
        uniqueUnits: 0,
        duplicateNopol: 0,
        invalidRows: 0
      },
      message: 'Sheet MASTER_UNIT belum ditemukan.'
    }, actor);
  }

  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
    return scopeMasterUnitResponse_({
      rows: [],
      meta: {
        sheetFound: true,
        headers: [],
        totalRows: 0,
        uniqueUnits: 0,
        duplicateNopol: 0,
        invalidRows: 0
      },
      message: 'Sheet MASTER_UNIT masih kosong.'
    }, actor);
  }

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0]
    .map(value => cleanText_(value));

  const fieldIndexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);

  if (fieldIndexes.nopol < 0) {
    throw new Error(
      'Header nomor polisi tidak ditemukan pada MASTER_UNIT. ' +
      'Gunakan salah satu header: ' +
      MASTER_UNIT_FIELD_ALIASES.nopol.join(', ') +
      '. Header terbaca: ' + headers.filter(Boolean).join(', ')
    );
  }

  const recommendedFieldLabels = {
    operatingPoint: 'Operating Point',
    toh: 'Nama TOH/WOH',
    statusUnit: 'Status Unit',
    kmOdometer: 'KM Odometer',
    status3L: 'Status 3L',
    status3LUpdatedAt: 'Timestamp Status 3L',
    status3LUpdatedBy: 'Update By Status 3L',
    calibrationNumber: 'Kalibrasi Ke',
    status3LNote: 'Catatan / Note'
  };
  const unmappedFields = Object.keys(recommendedFieldLabels)
    .filter(field => fieldIndexes[field] < 0)
    .map(field => recommendedFieldLabels[field]);

  // Satu kali pembacaan Sheet; hindari clone/konversi seluruh matriks ke string.
  // Hanya field yang benar-benar dipakai dashboard yang diformat di bawah.
  const usedFieldIndexes = Object.keys(fieldIndexes).map(key => fieldIndexes[key]).filter(index => index >= 0);
  const readColumnCount = Math.max(1, usedFieldIndexes.length ? Math.max.apply(null, usedFieldIndexes) + 1 : 1);
  const rawMasterValues = sheet.getLastRow() < 2 ? [] : sheet
    .getRange(2, 1, sheet.getLastRow() - 1, readColumnCount).getValues();

  const seenNopol = Object.create(null);
  const rows = [];
  let duplicateNopol = 0;
  let invalidRows = 0;
  let invalidNopolFormat = 0;

  // Pada database sehat, attribution yang sudah lengkap di MASTER_UNIT langsung
  // dipakai. Ledger histori yang dapat sangat besar hanya dibaca secara lazy bila
  // ada timestamp/actor yang belum lengkap dan memang memerlukan fallback.
  let status3LAuditMap = null;
  let status3LActivityMap = null;
  let status3LEvidenceLoaded = false;
  let scopedTotalRows = 0;
  rawMasterValues.forEach(row => {
    if (!canAccessOperatingPoint_(actor, masterUnitValue_(row, fieldIndexes.operatingPoint))) return;
    if (!row.some(value => cleanText_(value))) {
      return;
    }
    if (fieldIndexes.masterStatus >= 0 && normalizeLookupText_(row[fieldIndexes.masterStatus]) === 'INACTIVE') return;

    scopedTotalRows++;
    const nopol = masterUnitValue_(row, fieldIndexes.nopol);
    const nopolKey = normalizeNopol_(nopol);

    if (!nopolKey) {
      invalidRows++;
      return;
    }
    if (!isCanonicalNopol_(nopol)) invalidNopolFormat++;

    if (seenNopol[nopolKey]) {
      duplicateNopol++;
      return;
    }

    seenNopol[nopolKey] = true;


    const kmSnapshot = unitKmSnapshot_(row, fieldIndexes);
    const status3LEligibility = unitStatus3LEligibility_(kmSnapshot);
    const rawStatus3L = fieldIndexes.status3L >= 0 ? row[fieldIndexes.status3L] : '';
    const rawStatus3LUpdatedAt = fieldIndexes.status3LUpdatedAt >= 0 ? row[fieldIndexes.status3LUpdatedAt] : '';
    const rawStatus3LUpdatedBy = fieldIndexes.status3LUpdatedBy >= 0 ? row[fieldIndexes.status3LUpdatedBy] : '';
    const parsedStatus3LUpdatedAt = rawStatus3LUpdatedAt instanceof Date
      ? rawStatus3LUpdatedAt : parseDateTimeInput_(rawStatus3LUpdatedAt);
    const status3LTimestampReady = parsedStatus3LUpdatedAt instanceof Date && !isNaN(parsedStatus3LUpdatedAt.getTime());
    const status3LActorReady = normalizeLookupText_(rawStatus3L) !== 'DONE' ||
      isResolvedStatus3LActor_(rawStatus3LUpdatedBy);
    const needsStatus3LEvidence = Boolean(cleanText_(rawStatus3L)) &&
      (!status3LTimestampReady || !status3LActorReady);
    if (needsStatus3LEvidence && !status3LEvidenceLoaded) {
      status3LAuditMap = buildStatus3LAuditEvidenceMap_(ss.getSheetByName(UNIT_STATUS_3L_LOG_SHEET));
      status3LActivityMap = buildStatus3LActivityEvidenceMap_(ss.getSheetByName(KM_APP.SHEETS.LOG));
      status3LEvidenceLoaded = true;
    }
    const status3LAudit = needsStatus3LEvidence ? status3LAttributionEvidenceForRow_(
      status3LAuditMap, status3LActivityMap, row, fieldIndexes
    ) : null;
    rows.push({
      unitId: fieldIndexes.unitId >= 0 ? masterUnitValue_(row, fieldIndexes.unitId) : '',
      masterStatus: fieldIndexes.masterStatus >= 0 ? (masterUnitValue_(row, fieldIndexes.masterStatus) || 'ACTIVE') : 'ACTIVE',
      nopol: normalizeDisplayNopol_(nopol),
      operatingPoint: masterUnitValue_(row, fieldIndexes.operatingPoint),
      toh: masterUnitValue_(row, fieldIndexes.toh),
      statusUnit: masterUnitValue_(row, fieldIndexes.statusUnit),
      status3L: masterUnitValue_(row, fieldIndexes.status3L),
      status3LUpdatedAt: status3LAudit && status3LAudit.updatedAt ? status3LAudit.updatedAt : kmSnapshot.status3LUpdatedAt,
      status3LUpdatedBy: normalizeDoneStatus3LActor_(kmSnapshot.status3L,
        status3LAudit && status3LAudit.resolvedActor ? status3LAudit.updatedBy : kmSnapshot.status3LUpdatedBy),
      status3LNote: masterUnitValue_(row, fieldIndexes.status3LNote),
      kmOdometer: kmSnapshot.kmOdometer,
      kmDate: kmSnapshot.kmDate,
      kmUpdatedAt: kmSnapshot.kmUpdatedAt,
      calibrationNumber: kmSnapshot.calibrationNumber,
      kmRevision: kmSnapshot.kmRevision,
      kmPhotoId: kmSnapshot.kmPhotoId,
      calibrationDataComplete: status3LEligibility.dataComplete,
      status3LEligible: status3LEligibility.eligible,
      status3LEligibilityReason: status3LEligibility.reason
    });
  });

  rows.sort((a, b) =>
    a.nopol.localeCompare(b.nopol, 'id', { numeric: true })
  );

  const response = {
    rows: rows,
    unitStatuses: UNIT_MANUAL_OPERATIONAL_STATUSES.slice(),
    operationalStatuses: UNIT_OPERATIONAL_STATUSES.slice(),
    status3LOptions: getStatus3LOptions_(),
    meta: {
      sheetFound: true,
      headers: headers.filter(Boolean),
      totalRows: scopedTotalRows,
      uniqueUnits: rows.length,
      duplicateNopol: duplicateNopol,
      invalidRows: invalidRows,
      invalidNopolFormat: invalidNopolFormat,
      unmappedFields: unmappedFields,
      status3LReady: fieldIndexes.status3L >= 0,
      status3LEligibilityReady: fieldIndexes.kmOdometer >= 0 &&
        fieldIndexes.kmDate >= 0 && fieldIndexes.kmPhotoId >= 0 &&
        fieldIndexes.calibrationNumber >= 0,
      workflowColumnsReady: fieldIndexes.kmDate >= 0 &&
        fieldIndexes.kmUpdatedAt === fieldIndexes.kmDate + 1 &&
        fieldIndexes.kmUpdatedBy === fieldIndexes.kmUpdatedAt + 1 &&
        fieldIndexes.calibrationNumber === fieldIndexes.kmUpdatedBy + 1 &&
        fieldIndexes.status3L === fieldIndexes.calibrationNumber + 1 &&
        fieldIndexes.status3LUpdatedAt === fieldIndexes.status3L + 1 &&
        fieldIndexes.status3LUpdatedBy === fieldIndexes.status3LUpdatedAt + 1 &&
        fieldIndexes.status3LNote === fieldIndexes.status3LUpdatedBy + 1,
      statusValidationReady: isMasterUnitStatusValidationReady_(sheet, fieldIndexes.statusUnit),
      editTriggerReady: isKalibrasiKmEditTriggerReady_(),
      operatingPointReady: fieldIndexes.operatingPoint >= 0,
      schemaReady: fieldIndexes.nopol >= 0 && fieldIndexes.operatingPoint >= 0 &&
        fieldIndexes.statusUnit >= 0 && fieldIndexes.kmOdometer >= 0 && fieldIndexes.status3L >= 0 &&
        fieldIndexes.calibrationNumber >= 0 && fieldIndexes.status3LUpdatedAt >= 0 &&
        fieldIndexes.status3LUpdatedBy >= 0 &&
        fieldIndexes.status3LNote >= 0
    }
  };

  const scopedResponse = scopeMasterUnitResponse_(response, actor, true);
  putCachedJson_(responseCacheKey, scopedResponse, 45);
  return Object.assign({}, scopedResponse, {
    submissionSchedule: getCalibrationSubmissionSchedule_()
  });
}

function scopeMasterUnitResponse_(response, actor, alreadyScoped) {
  const scopedRows = alreadyScoped
    ? (response.rows || [])
    : filterRowsByActorScope_(response.rows || [], actor);
  const scoped = Object.assign({}, response, {
    rows: scopedRows,
    submissionSchedule: getCalibrationSubmissionSchedule_()
  });
  const baseMeta = Object.assign({}, response.meta || {});
  scoped.meta = Object.assign(baseMeta, {
    uniqueUnits: scopedRows.length,
    scoped: !actor.globalScope,
    visibleOperatingPoints: actor.operatingPoints.slice()
  });
  return scoped;
}

function findMasterUnitRowNumberByNopol_(sheet, indexes, nopol) {
  if (!sheet || sheet.getLastRow() < 2 || !indexes || indexes.nopol < 0) return 0;
  const target = normalizeNopol_(nopol);
  if (!target) return 0;
  const range = sheet.getRange(2, indexes.nopol + 1, sheet.getLastRow() - 1, 1);
  try {
    const pattern = normalizedNopolTextFinderPattern_(target);
    const matches = range.createTextFinder(pattern)
      .useRegularExpression(true)
      .matchCase(false)
      .matchEntireCell(true)
      .findAll();
    if (matches.length > 1) throw new Error('Nomor polisi duplikat di MASTER_UNIT.');
    if (matches.length === 1) return matches[0].getRow();
    return 0;
  } catch (error) {
    if (/duplikat/i.test(String(error && error.message || error))) throw error;
    console.warn('TextFinder MASTER_UNIT fallback ke scan kolom:', error && error.message ? error.message : error);
    const plates = range.getDisplayValues();
    const matches = [];
    plates.forEach((row, offset) => {
      if (normalizeNopol_(row[0]) === target) matches.push(offset + 2);
    });
    if (matches.length > 1) throw new Error('Nomor polisi duplikat di MASTER_UNIT.');
    return matches.length ? matches[0] : 0;
  }
}

function assertUnitScope_(actor, nopol) {
  if (actor.globalScope) return;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('MASTER_UNIT belum tersedia.');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  if (indexes.nopol < 0 || indexes.operatingPoint < 0) {
    throw new Error('Header Nomor Polisi / Operating Point belum tersedia pada MASTER_UNIT.');
  }
  const rowNumber = findMasterUnitRowNumberByNopol_(sheet, indexes, nopol);
  if (!rowNumber) throw new Error('Nomor polisi tidak ditemukan.');
  const operatingPoint = sheet.getRange(rowNumber, indexes.operatingPoint + 1).getDisplayValue();
  if (!canAccessOperatingPoint_(actor, operatingPoint)) {
    throw new Error('Akses ditolak: unit berada di luar scope Operating Point akun Anda.');
  }
}


// ============================================================================
// 02. AUTHENTICATION, SESSIONS & RBAC
// ============================================================================
/** Application configuration, authentication, sessions, users, roles, and RBAC. */
const KM_APP = Object.freeze({
  SHEETS: {
    CONFIG: 'MASTER_CONFIG',
    LOG: 'KM_ACTIVITY_LOG',
    MASTER_UNIT: 'MASTER_UNIT',
    USERS: 'KM_USERS',
    ROLES: 'KM_ROLES',
    SESSIONS: 'KM_SESSIONS',
    AUTH_LOG: 'KM_AUTH_LOG'
  },


  LOG_HEADERS: [
    'LOG ID',
    'TIMESTAMP',
    'NOMOR POLISI',
    'USER EMAIL',
    'ACTION',
    'FIELD NAME',
    'OLD VALUE',
    'NEW VALUE',
    'NOTE',
    'USER ROLE',
    'SOURCE',
    'DETAILS JSON'
  ],

  USER_HEADERS: [
    // Kolom existing dipertahankan agar migrasi bersifat append-only.
    'EMAIL',
    'ROLES',
    'OPERATING POINTS',
    'ACTIVE',
    'DISPLAY NAME',
    'CREATED AT',
    'CREATED BY',
    'UPDATED AT',
    'UPDATED BY',
    // Kolom autentikasi enterprise. Jangan pernah menampilkan hash/salt ke frontend.
    'USER ID',
    'USERNAME',
    'PASSWORD HASH',
    'PASSWORD SALT',
    'PASSWORD ITERATIONS',
    'PASSWORD CHANGED AT',
    'MUST CHANGE PASSWORD',
    'FAILED LOGIN COUNT',
    'LOCKED UNTIL',
    'LAST LOGIN AT',
    'SESSION VERSION',
    'PASSWORD ALGORITHM'
  ],



  CONFIG: {}
});

const KM_RBAC = Object.freeze({
  ROLES: [
    'SUPER_ADMIN',
    'ADMIN_BITRANS',
    'OH',
    'IT',
    'TEAM_3L',
    'OPC',
    'VIEWER'
  ],
  PERMISSIONS: {
    ACCESS_APP: ['SUPER_ADMIN', 'ADMIN_BITRANS', 'OH', 'IT', 'TEAM_3L', 'OPC', 'VIEWER'],
    VIEW_MONITORING: ['SUPER_ADMIN', 'ADMIN_BITRANS', 'OH', 'IT', 'TEAM_3L', 'OPC', 'VIEWER'],
    VIEW_MASTER_UNIT: ['SUPER_ADMIN', 'ADMIN_BITRANS', 'OH', 'IT', 'TEAM_3L', 'OPC', 'VIEWER'],
    UPDATE_UNIT: ['SUPER_ADMIN', 'ADMIN_BITRANS', 'OPC'],
    UPDATE_STATUS_3L: ['SUPER_ADMIN', 'ADMIN_BITRANS', 'TEAM_3L'],
    MANAGE_MASTER_UNIT: ['SUPER_ADMIN', 'ADMIN_BITRANS'],
    REQUEST_MASTER_UNIT_CHANGE: ['SUPER_ADMIN', 'ADMIN_BITRANS', 'OH', 'OPC'],
    APPROVE_MASTER_UNIT_CHANGE: ['SUPER_ADMIN', 'ADMIN_BITRANS'],
    MANAGE_USERS: ['SUPER_ADMIN'],
    ADMIN_SETUP: ['SUPER_ADMIN'],
    VIEW_AUDIT: ['SUPER_ADMIN', 'ADMIN_BITRANS']
  },
  GLOBAL_SCOPE_ROLES: [],
  SOURCES: ['WEB_APP', 'SHEET']
});

const KM_ROLE_HEADERS = Object.freeze([
  'ROLE', 'DISPLAY NAME', 'PERMISSIONS', 'GLOBAL SCOPE', 'ACTIVE', 'SYSTEM ROLE',
  'CREATED AT', 'CREATED BY', 'UPDATED AT', 'UPDATED BY'
]);

const KM_PERMISSION_LABELS = Object.freeze({
  ACCESS_APP: 'Akses aplikasi', VIEW_MONITORING: 'Lihat Monitoring',
  VIEW_MASTER_UNIT: 'Lihat Pengajuan Kalibrasi', UPDATE_UNIT: 'Update status & odometer',
  UPDATE_STATUS_3L: 'Update Status 3L',
  MANAGE_MASTER_UNIT: 'Kelola master unit',
  REQUEST_MASTER_UNIT_CHANGE: 'Ajukan perubahan master unit',
  APPROVE_MASTER_UNIT_CHANGE: 'Approve perubahan master unit',
  MANAGE_USERS: 'Kelola user & role', ADMIN_SETUP: 'Setup database', VIEW_AUDIT: 'Lihat audit trail'
});

const KM_AUTH = Object.freeze({
  SESSION_TTL_HOURS: 8,
  IDLE_TIMEOUT_MINUTES: 30,
  SESSION_TOUCH_MINUTES: 10,
  SESSION_CACHE_SECONDS: 600,
  USER_DIRECTORY_CACHE_SECONDS: 600,
  UNKNOWN_LOGIN_CACHE_SECONDS: 900,
  MAX_UNKNOWN_LOGIN_ATTEMPTS: 8,
  MAX_FAILED_LOGINS: 5,
  LOCK_MINUTES: 15,
  PASSWORD_MIN_LENGTH: 10,
  PASSWORD_MAX_LENGTH: 128,
  PASSWORD_ITERATIONS: 10000,
  TEMP_PASSWORD_LENGTH: 18,
  PEPPER_PROPERTY: 'KM_AUTH_PEPPER_V1',
  SESSION_CACHE_PREFIX: 'km-auth-session-v1:',
  SESSION_REVOKED_CACHE_PREFIX: 'km-auth-revoked-v1:',
  UNKNOWN_LOGIN_PREFIX: 'km-auth-unknown-v1:',
  SESSION_HEADERS: [
    'SESSION ID',
    'USER ID',
    'TOKEN HASH',
    'CREATED AT',
    'EXPIRES AT',
    'LAST ACTIVITY AT',
    'STATUS',
    'SESSION VERSION',
    'USER AGENT HASH',
    'REVOKED AT',
    'REVOKE REASON'
  ],
  AUTH_LOG_HEADERS: [
    'EVENT ID',
    'TIMESTAMP',
    'IDENTIFIER',
    'USER ID',
    'EMAIL',
    'ACTION',
    'STATUS',
    'NOTE',
    'USER AGENT HASH',
    'SESSION ID'
  ]
});

const KM_SHEET_EDIT_TRIGGER = Object.freeze({
  HANDLER: 'handleKalibrasiKmSheetEdit',
  ENABLED_PROPERTY: 'KM_INSTALLABLE_EDIT_TRIGGER_ENABLED_V1',
  ID_PROPERTY: 'KM_INSTALLABLE_EDIT_TRIGGER_ID_V2',
  LAST_FIRED_PROPERTY: 'KM_INSTALLABLE_EDIT_TRIGGER_LAST_FIRED_AT_V1'
});

// Fallback operasional bila Google tidak mengekspos email editor pada edit langsung Sheet.
// Marker teknis lama tidak boleh lagi ditulis atau ditampilkan ke user.
const KM_STATUS_3L_SHEET_FALLBACK = 'Team 3L';
const KM_STATUS_3L_AUTOMATION_ACTOR = 'Team 3L External Sync';
const KM_STATUS_3L_LEGACY_UNRESOLVED = 'SHEET_EDITOR_UNRESOLVED';

const KM_STATUS_3L_WATCH = Object.freeze({
  HANDLER: 'syncExternalStatus3LChanges',
  ENABLED_PROPERTY: 'KM_STATUS_3L_WATCH_ENABLED_V1',
  ID_PROPERTY: 'KM_STATUS_3L_WATCH_TRIGGER_ID_V1',
  LAST_SIGNATURE_PROPERTY: 'KM_STATUS_3L_WATCH_SIGNATURE_V1',
  LAST_RUN_PROPERTY: 'KM_STATUS_3L_WATCH_LAST_RUN_AT_V1',
  LAST_SUCCESS_PROPERTY: 'KM_STATUS_3L_WATCH_LAST_SUCCESS_AT_V2',
  LAST_ERROR_PROPERTY: 'KM_STATUS_3L_WATCH_LAST_ERROR_V2',
  ERROR_COUNT_PROPERTY: 'KM_STATUS_3L_WATCH_ERROR_COUNT_V2',
  INTERVAL_MINUTES: 1,
  MANUAL_EDIT_GRACE_MS: 2 * 60 * 1000
});

// Request-scoped actor. Set hanya setelah token session / actor Sheet berhasil diverifikasi.
let KM_REQUEST_ACTOR_ = null;
// Memo per eksekusi Apps Script. Ini menghindari parsing Sheet/cache yang sama
// berulang kali ketika satu endpoint membangun actor, permission, dan bootstrap.
let KM_REQUEST_ROLE_DIRECTORY_ = null;
let KM_REQUEST_USER_DIRECTORY_ = null;
let KM_REQUEST_ACTIVE_ROLE_NAMES_ = null;
let KM_REQUEST_CONFIG_COVERAGE_ = null;
let KM_REQUEST_UNIT_COVERAGE_ = null;
let KM_REQUEST_EFFECTIVE_COVERAGE_ = null;

function acquireScriptLock_(timeoutMs, operationLabel) {
  const lock = LockService.getScriptLock();
  const timeout = Math.max(1000, Number(timeoutMs) || 20000);
  try {
    if (lock.tryLock(timeout)) return lock;
  } catch (error) {
    console.warn('Lock gagal untuk ' + (operationLabel || 'operasi') + ': ' +
      String(error && error.message ? error.message : error));
  }
  throw new Error('SYSTEM_BUSY: Transaksi lain belum selesai. ' +
    'Permintaan ini aman untuk dicoba ulang.');
}


const MASTER_UNIT_FIELD_ALIASES = Object.freeze({
  nopol: [
    'NOMOR POLISI',
    'NO POLISI',
    'NO. POLISI',
    'NOMOR POL',
    'NO POL',
    'PLAT NOMOR',
    'NOPOL',
    'VEHICLE NUMBER',
    'VEHICLE NO',
    'POLICE NUMBER',
    'LICENSE PLATE'
  ],
  unitType: [
    'JENIS UNIT',
    'TIPE UNIT',
    'TIPE',
    'TYPE UNIT',
    'UNIT TYPE',
    'JENIS KENDARAAN'
  ],
  operatingPoint: [
    'OPERATING POINT',
    'OPERATIONAL POINT',
    'OP'
  ],
  toh: ['NAMA TOH/WOH', 'TOH/WOH', 'NAMA TOH', 'NAMA WOH'],
  kmOdometer: ['KM ODOMETER', 'ODOMETER'],
  kmPhotoId: ['FOTO ODOMETER ID'],
  kmPhotoAddedAt: ['FOTO DITAMBAHKAN AT'],
  kmPhotoAddedBy: ['FOTO DIUNGGAH OLEH'],
  kmDate: ['TANGGAL UPDATE KM', 'TANGGAL KM'],
  kmUpdatedAt: ['KM UPDATED AT'],
  kmUpdatedBy: ['KM UPDATED BY'],
  calibrationNumber: ['KALIBRASI KE', 'KALIBRASI_KM_KE', 'CALIBRATION NUMBER'],
  status3L: ['STATUS 3L', 'STATUS_3L', 'STATUS THREE L'],
  status3LUpdatedAt: ['STATUS 3L UPDATED AT', 'TIMESTAMP STATUS 3L', 'WAKTU UPDATE STATUS 3L'],
  status3LUpdatedBy: ['STATUS 3L UPDATED BY', 'STATUS 3L UPDATE BY', 'STATUS 3L DIUPDATE OLEH'],
  status3LNote: ['CATATAN / NOTE', 'CATATAN STATUS 3L', 'CATATAN 3L', 'STATUS 3L NOTE'],
  statusUnit: [
    'STATUS UNIT',
    'STATUS KENDARAAN',
    'STATUS AKTIF',
    'STATUS'
  ],
  unitId: ['UNIT ID', 'UNIT_ID'],
  masterStatus: ['MASTER STATUS', 'MASTER_STATUS', 'STATUS MASTER'],
  masterCreatedAt: ['MASTER CREATED AT'],
  masterCreatedBy: ['MASTER CREATED BY'],
  masterUpdatedAt: ['MASTER UPDATED AT'],
  masterUpdatedBy: ['MASTER UPDATED BY'],
  lastMovementId: ['LAST MOVEMENT ID']
});

function normalizeEmail_(value) {
  return cleanText_(value).toLowerCase();
}

function normalizeUsername_(value) {
  return cleanText_(value).toLowerCase().replace(/\s+/g, '');
}

function splitAccessValues_(value) {
  const source = Array.isArray(value) ? value : String(value || '').split(/[;,\n]/);
  return [...new Set(source.map(item => cleanText_(item)).filter(Boolean))];
}

const KM_ROLE_ALIASES = Object.freeze({
  // Legacy role code dipertahankan sebagai alias agar user/session/database lama
  // tetap kompatibel setelah nama role digeneralisasi menjadi OH.
  APPROVER_OH: 'OH'
});

function canonicalRoleCode_(value) {
  const normalized = normalizeLookupText_(value).replace(/[\s-]+/g, '_');
  return KM_ROLE_ALIASES[normalized] || normalized;
}

function defaultRoleDirectory_() {
  return KM_RBAC.ROLES.map(role => ({
    role: role,
    displayName: role.replace(/_/g, ' '),
    permissions: Object.keys(KM_RBAC.PERMISSIONS).filter(permission =>
      KM_RBAC.PERMISSIONS[permission].includes(role)),
    globalScope: KM_RBAC.GLOBAL_SCOPE_ROLES.includes(role),
    active: true,
    systemRole: true
  }));
}

function readRoleDirectory_(forceRefresh) {
  if (forceRefresh) KM_REQUEST_ACTIVE_ROLE_NAMES_ = null;
  if (!forceRefresh && KM_REQUEST_ROLE_DIRECTORY_) return KM_REQUEST_ROLE_DIRECTORY_;
  const cached = !forceRefresh && getCachedJson_('roles');
  if (cached && cached.length) {
    // Cache dari deployment lama dapat masih menyimpan APPROVER_OH.
    // Normalisasi saat read agar rename ke OH langsung konsisten tanpa menunggu TTL cache habis.
    const normalizedCached = cached.map(item => {
      const role = canonicalRoleCode_(item && item.role);
      return Object.assign({}, item, {
        role: role,
        displayName: role === 'OH' && item && item.systemRole ? 'OH' : (item && item.displayName || role)
      });
    });
    KM_REQUEST_ROLE_DIRECTORY_ = normalizedCached;
    return normalizedCached;
  }
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.ROLES);
    if (!sheet || sheet.getLastRow() < 2) {
      KM_REQUEST_ROLE_DIRECTORY_ = defaultRoleDirectory_();
      return KM_REQUEST_ROLE_DIRECTORY_;
    }
    const map = getHeaderMap_(sheet);
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const allowedPermissions = new Set(Object.keys(KM_RBAC.PERMISSIONS));
    const roles = values.map((row, index) => ({
      rowNumber: index + 2,
      role: canonicalRoleCode_(getRowValue_(row, map, 'ROLE')),
      displayName: canonicalRoleCode_(getRowValue_(row, map, 'ROLE')) === 'OH' &&
        boolValue_(getRowValue_(row, map, 'SYSTEM ROLE'))
          ? 'OH' : cleanText_(getRowValue_(row, map, 'DISPLAY NAME')),
      permissions: splitAccessValues_(getRowValue_(row, map, 'PERMISSIONS'))
        .map(normalizeLookupText_).filter(permission => allowedPermissions.has(permission)),
      globalScope: false,
      active: boolValue_(getRowValue_(row, map, 'ACTIVE')),
      systemRole: boolValue_(getRowValue_(row, map, 'SYSTEM ROLE'))
    })).filter(item => item.role);
    if (!roles.some(item => item.role === 'SUPER_ADMIN')) {
      throw new Error('KM_ROLES tidak memiliki SUPER_ADMIN. Jalankan setup aplikasi untuk memperbaiki konfigurasi.');
    }
    putCachedJson_('roles', roles, KM_AUTH.USER_DIRECTORY_CACHE_SECONDS);
    KM_REQUEST_ROLE_DIRECTORY_ = roles;
    return roles;
  } catch (error) {
    console.warn('Directory role tidak dapat dibaca dengan aman: ' + (error && error.message ? error.message : error));
    KM_REQUEST_ROLE_DIRECTORY_ = [];
    return [];
  }
}

function normalizeRoles_(value) {
  if (!KM_REQUEST_ACTIVE_ROLE_NAMES_) {
    KM_REQUEST_ACTIVE_ROLE_NAMES_ = new Set(
      readRoleDirectory_().filter(role => role.active).map(role => role.role)
    );
  }
  return splitAccessValues_(value)
    .map(canonicalRoleCode_)
    .filter(role => KM_REQUEST_ACTIVE_ROLE_NAMES_.has(role));
}

function isActiveAccessValue_(value) {
  return ['TRUE', 'YA', 'YES', '1', 'ACTIVE', 'AKTIF'].includes(normalizeLookupText_(value));
}

function boolValue_(value) {
  return value === true || isActiveAccessValue_(value);
}

function getSpreadsheetOwnerEmail_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const owner = ss && typeof ss.getOwner === 'function' ? ss.getOwner() : null;
    return normalizeEmail_(owner && typeof owner.getEmail === 'function' ? owner.getEmail() : '');
  } catch (error) {
    return '';
  }
}

function getSheetUserEmail_(event) {
  if (event) {
    try {
      if (event.user && typeof event.user.getEmail === 'function') {
        const eventEmail = normalizeEmail_(event.user.getEmail());
        if (eventEmail) return eventEmail;
      }
    } catch (error) {}
  }
  try {
    // ActiveUser adalah user aktif; EffectiveUser adalah otorisasi eksekusi.
    // Keduanya boleh sama ketika pembuat trigger sendiri mengedit Sheet.
    // Jangan pernah menggunakan EffectiveUser sebagai pengganti editor.
    return normalizeEmail_(Session.getActiveUser().getEmail() || '');
  } catch (error) {
    return '';
  }
}

function status3LActorDisplay_(value) {
  const actor = cleanText_(value);
  if (normalizeLookupText_(actor) === normalizeLookupText_(KM_STATUS_3L_LEGACY_UNRESOLVED)) {
    return KM_STATUS_3L_SHEET_FALLBACK;
  }
  return actor;
}

function normalizeDoneStatus3LActor_(status3L, actor) {
  const display = status3LActorDisplay_(actor);
  if (normalizeLookupText_(status3L) !== 'DONE') return display;
  const key = normalizeLookupText_(display);
  if (display && key !== 'SYSTEM' && key !== normalizeLookupText_(KM_STATUS_3L_LEGACY_UNRESOLVED)) {
    return display;
  }
  return KM_STATUS_3L_SHEET_FALLBACK;
}

function preferredDoneStatus3LActor_(primaryActor, evidenceActor) {
  const candidates = [primaryActor, evidenceActor];
  for (let index = 0; index < candidates.length; index++) {
    const display = status3LActorDisplay_(candidates[index]);
    const key = normalizeLookupText_(display);
    if (display && key !== 'SYSTEM' &&
        key !== normalizeLookupText_(KM_STATUS_3L_LEGACY_UNRESOLVED) &&
        isResolvedStatus3LActor_(display)) {
      return display;
    }
  }
  return KM_STATUS_3L_SHEET_FALLBACK;
}

function resolveManualStatus3LEditorEmail_(event, actor) {
  // Status 3L manual tidak boleh pernah diatribusikan sebagai SYSTEM. Ambil email editor
  // dari event.user/ActiveUser terlebih dahulu, lalu actor terverifikasi. Bila Google
  // menyembunyikan identitas, gunakan label operasional "Team 3L" agar dashboard/audit
  // tidak menampilkan marker teknis SHEET_EDITOR_UNRESOLVED.
  const candidates = [
    getSheetUserEmail_(event),
    actor && actor.email
  ];
  for (let i = 0; i < candidates.length; i++) {
    const raw = status3LActorDisplay_(candidates[i]);
    const email = normalizeEmail_(raw);
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return email;
    if (normalizeLookupText_(raw) === normalizeLookupText_(KM_STATUS_3L_SHEET_FALLBACK)) {
      return KM_STATUS_3L_SHEET_FALLBACK;
    }
  }
  return KM_STATUS_3L_SHEET_FALLBACK;
}

function readUserDirectory_(forceRefresh) {
  if (!forceRefresh && KM_REQUEST_USER_DIRECTORY_) return KM_REQUEST_USER_DIRECTORY_;
  const cached = !forceRefresh && getCachedJson_('users');
  if (cached) {
    // Normalize role aliases from cache generated by an older deployment.
    const normalizedCached = cached.map(user => Object.assign({}, user, {
      roles: [...new Set((user && user.roles || []).map(canonicalRoleCode_).filter(Boolean))]
    }));
    KM_REQUEST_USER_DIRECTORY_ = normalizedCached;
    return normalizedCached;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(KM_APP.SHEETS.USERS);
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) {
    KM_REQUEST_USER_DIRECTORY_ = [];
    return KM_REQUEST_USER_DIRECTORY_;
  }

  const map = getHeaderMap_(sheet);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const users = values.map((row, index) => userRecordFromRow_(row, map, index + 2))
    .filter(user => user.email || user.username);

  putCachedJson_('users', users, KM_AUTH.USER_DIRECTORY_CACHE_SECONDS);
  KM_REQUEST_USER_DIRECTORY_ = users;
  return users;
}

function userRecordFromRow_(row, map, rowNumber) {
  return {
    rowNumber: rowNumber,
    userId: cleanText_(getRowValue_(row, map, 'USER ID')),
    email: normalizeEmail_(getRowValue_(row, map, 'EMAIL')),
    username: normalizeUsername_(getRowValue_(row, map, 'USERNAME')),
    roles: normalizeRoles_(getRowValue_(row, map, 'ROLES')),
    operatingPoints: splitAccessValues_(getRowValue_(row, map, 'OPERATING POINTS')),
    active: isActiveAccessValue_(getRowValue_(row, map, 'ACTIVE')),
    displayName: cleanText_(getRowValue_(row, map, 'DISPLAY NAME')),
    passwordHash: cleanText_(getRowValue_(row, map, 'PASSWORD HASH')),
    passwordSalt: cleanText_(getRowValue_(row, map, 'PASSWORD SALT')),
    passwordIterations: Number(getRowValue_(row, map, 'PASSWORD ITERATIONS')) || KM_AUTH.PASSWORD_ITERATIONS,
    passwordAlgorithm: cleanText_(getRowValue_(row, map, 'PASSWORD ALGORITHM')) || 'PBKDF2_SHA256_V1',
    passwordChangedAt: formatDateTime_(getRowValue_(row, map, 'PASSWORD CHANGED AT')),
    mustChangePassword: boolValue_(getRowValue_(row, map, 'MUST CHANGE PASSWORD')),
    failedLoginCount: Number(getRowValue_(row, map, 'FAILED LOGIN COUNT')) || 0,
    lockedUntilRaw: getRowValue_(row, map, 'LOCKED UNTIL'),
    lockedUntil: formatDateTime_(getRowValue_(row, map, 'LOCKED UNTIL')),
    lastLoginAt: formatDateTime_(getRowValue_(row, map, 'LAST LOGIN AT')),
    sessionVersion: Math.max(1, Number(getRowValue_(row, map, 'SESSION VERSION')) || 1),
    createdAt: formatDateTime_(getRowValue_(row, map, 'CREATED AT')),
    createdBy: cleanText_(getRowValue_(row, map, 'CREATED BY')),
    updatedAt: formatDateTime_(getRowValue_(row, map, 'UPDATED AT')),
    updatedBy: cleanText_(getRowValue_(row, map, 'UPDATED BY'))
  };
}

function publicUserRecord_(user) {
  const lockDate = user && user.lockedUntilRaw instanceof Date
    ? user.lockedUntilRaw
    : parseDateTimeInput_(user && user.lockedUntilRaw);
  const activeLockedUntil = lockDate instanceof Date && lockDate.getTime() > new Date().getTime()
    ? formatDateTime_(lockDate)
    : '';
  return {
    rowNumber: user.rowNumber,
    userId: user.userId || '',
    email: user.email || '',
    username: user.username || '',
    roles: (user.roles || []).slice(),
    operatingPoints: (user.operatingPoints || []).slice(),
    active: Boolean(user.active),
    displayName: user.displayName || '',
    passwordConfigured: Boolean(user.passwordHash && user.passwordSalt),
    mustChangePassword: Boolean(user.mustChangePassword),
    failedLoginCount: Number(user.failedLoginCount) || 0,
    lockedUntil: activeLockedUntil,
    lastLoginAt: user.lastLoginAt || '',
    passwordChangedAt: user.passwordChangedAt || '',
    createdAt: user.createdAt || '',
    createdBy: user.createdBy || '',
    updatedAt: user.updatedAt || '',
    updatedBy: user.updatedBy || ''
  };
}

function permissionsForRoles_(roles) {
  const selected = new Set(roles || []);
  const permissions = [];
  const directory = readRoleDirectory_();

  directory
    .filter(role => role.active && selected.has(role.role))
    .forEach(role => {
      (role.permissions || []).forEach(permission => permissions.push(permission));
    });

  // KM_ROLES adalah source of truth setelah role dibuat. Mapping pada KM_RBAC
  // hanya menjadi default saat database belum ada atau role pertama kali dibuat;
  // jangan menambahkan ulang permission yang sengaja dihapus oleh administrator.
  return [...new Set(permissions)]
    .filter(permission => Object.prototype.hasOwnProperty.call(KM_RBAC.PERMISSIONS, permission));
}

function buildActorFromUser_(user, source, sessionMeta) {
  if (!user || !user.active) throw new Error('AUTH_REQUIRED: akun tidak aktif atau tidak ditemukan.');
  const roles = normalizeRoles_(user.roles);
  if (!roles.length) throw new Error('AUTH_REQUIRED: akun belum memiliki role yang valid.');
  const operatingPoints = splitAccessValues_(user.operatingPoints);
  const globalScope = operatingPoints.includes('*');
  const scopedOperatingPoints = globalScope ? ['*'] : operatingPoints;
  return {
    authorized: true,
    userId: user.userId || '',
    email: user.email || '',
    username: user.username || '',
    displayName: user.displayName || '',
    roles,
    primaryRole: roles[0],
    operatingPoints: scopedOperatingPoints,
    operatingPointKeys: globalScope ? [] : scopedOperatingPoints.map(normalizeLookupText_),
    globalScope,
    permissions: permissionsForRoles_(roles),
    mustChangePassword: Boolean(user.mustChangePassword),
    source: KM_RBAC.SOURCES.includes(source) ? source : 'WEB_APP',
    sessionId: sessionMeta && sessionMeta.sessionId || '',
    sessionExpiresAt: sessionMeta && sessionMeta.expiresAt || ''
  };
}

function getCurrentUserContext_(options) {
  const settings = options || {};
  const source = KM_RBAC.SOURCES.includes(settings.source) ? settings.source : 'WEB_APP';

  // Nested internal calls tanpa token/source eksplisit boleh reuse actor request yang sudah diverifikasi.
  // Endpoint publik yang membawa token selalu divalidasi ulang sehingga tidak pernah percaya global state lama.
  if (!settings.forceRevalidate && !settings.token && !settings.source && KM_REQUEST_ACTOR_ && KM_REQUEST_ACTOR_.authorized) {
    return KM_REQUEST_ACTOR_;
  }

  if (source === 'SHEET') {
    const email = getSheetUserEmail_(settings.event);
    // Sheet edit authorization must not use a 10-minute stale RBAC cache. A stale
    // role/user cache previously made a legitimate TEAM_3L editor look unauthorized
    // and the trigger then wrote e.oldValue back to STATUS 3L.
    readRoleDirectory_(true);
    const user = readUserDirectory_(true).find(item => item.email === email && item.active);
    if (!email || !user) {
      const unauthorized = {
        authorized: false, userId: '', email: email || '', username: '', displayName: '',
        roles: [], operatingPoints: [], globalScope: false, permissions: [],
        mustChangePassword: false, source: 'SHEET'
      };
      if (settings.allowUnauthorized) return unauthorized;
      throw new Error('AUTH_REQUIRED: akun Google editor Sheet belum terdaftar/aktif di KM_USERS.');
    }
    const actor = buildActorFromUser_(user, 'SHEET');
    KM_REQUEST_ACTOR_ = actor;
    return actor;
  }

  if (!settings.token) {
    if (settings.allowUnauthorized) {
      return {
        authorized: false, userId: '', email: '', username: '', displayName: '', roles: [],
        operatingPoints: [], globalScope: false, permissions: [], mustChangePassword: false,
        source: 'WEB_APP'
      };
    }
    throw new Error('AUTH_REQUIRED: silakan login kembali.');
  }

  const actor = validateSessionToken_(settings.token, settings);
  KM_REQUEST_ACTOR_ = actor;
  return actor;
}

function requirePermission_(permission, operatingPoint, options) {
  const actor = getCurrentUserContext_(options);
  if (actor.mustChangePassword && !(options && options.allowPasswordChange)) {
    throw new Error('PASSWORD_CHANGE_REQUIRED: password sementara harus diganti sebelum menggunakan aplikasi.');
  }
  if (!actor.permissions.includes(permission)) {
    throw new Error('ACCESS_DENIED: role ' + actor.roles.join(', ') + ' tidak memiliki izin ' + permission + '.');
  }
  if (operatingPoint && !canAccessOperatingPoint_(actor, operatingPoint)) {
    throw new Error('ACCESS_DENIED: Operating Point ' + operatingPoint + ' berada di luar scope akun Anda.');
  }
  return actor;
}

function canAccessOperatingPoint_(actor, operatingPoint) {
  if (!actor || !actor.authorized) return false;
  if (actor.globalScope) return true;
  const target = normalizeLookupText_(operatingPoint);
  const scopes = Array.isArray(actor.operatingPointKeys)
    ? actor.operatingPointKeys
    : (actor.operatingPoints || []).map(normalizeLookupText_);
  return scopes.includes(target);
}

function filterRowsByActorScope_(rows, actor) {
  return actor.globalScope ? rows.slice() : rows.filter(row =>
    canAccessOperatingPoint_(actor, row.operatingPoint)
  );
}

function publicActorContext_(actor) {
  return {
    authorized: Boolean(actor.authorized),
    userId: actor.userId || '',
    email: actor.email || '',
    username: actor.username || '',
    displayName: actor.displayName || '',
    roles: (actor.roles || []).slice(),
    primaryRole: actor.primaryRole || '',
    operatingPoints: (actor.operatingPoints || []).slice(),
    globalScope: Boolean(actor.globalScope),
    permissions: (actor.permissions || []).slice(),
    mustChangePassword: Boolean(actor.mustChangePassword),
    sessionExpiresAt: actor.sessionExpiresAt || ''
  };
}

function migrateLegacyUserRoleAliases_(usersSheet) {
  if (!usersSheet || usersSheet.getLastRow() < 2) return 0;
  const map = getHeaderMap_(usersSheet);
  if (!map['ROLES']) return 0;
  const rowCount = usersSheet.getLastRow() - 1;
  const range = usersSheet.getRange(2, map['ROLES'], rowCount, 1);
  const values = range.getDisplayValues();
  let changed = 0;
  const normalized = values.map(row => {
    const original = splitAccessValues_(row[0]);
    const canonical = [...new Set(original.map(canonicalRoleCode_).filter(Boolean))];
    const next = canonical.join(', ');
    if (next !== cleanText_(row[0])) changed++;
    return [next];
  });
  if (changed) range.setValues(normalized);
  return changed;
}

function setupAuthDatabase_(ss) {
  const users = getOrCreateSheet_(ss, KM_APP.SHEETS.USERS);
  const roles = setupRoleDatabase_(ss);
  const sessions = getOrCreateSheet_(ss, KM_APP.SHEETS.SESSIONS);
  const authLog = getOrCreateSheet_(ss, KM_APP.SHEETS.AUTH_LOG);
  ensureHeaders_(users, KM_APP.USER_HEADERS);
  migrateLegacyUserRoleAliases_(users);
  ensureHeaders_(sessions, KM_AUTH.SESSION_HEADERS);
  ensureHeaders_(authLog, KM_AUTH.AUTH_LOG_HEADERS);
  setupLogSheet_(getOrCreateSheet_(ss, KM_APP.SHEETS.LOG));

  [users, roles, sessions, authLog].forEach(sheet => runNonCritical_('freeze auth header', () => sheet.setFrozenRows(1)));
  runNonCritical_('hapus proteksi otomatis lama', removeAutomaticSheetProtections_);
  invalidateSecurityCaches_();
  return { users, roles, sessions, authLog };
}

function setupRoleDatabase_(ss) {
  const sheet = getOrCreateSheet_(ss, KM_APP.SHEETS.ROLES);
  ensureHeaders_(sheet, KM_ROLE_HEADERS);
  const map = getHeaderMap_(sheet);
  const existing = sheet.getLastRow() < 2 ? [] :
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const known = new Set(existing.map(row => canonicalRoleCode_(getRowValue_(row, map, 'ROLE'))));
  const now = new Date();
  const owner = getSpreadsheetOwnerEmail_();
  const rows = defaultRoleDirectory_().filter(role => !known.has(role.role)).map(role => {
    const row = new Array(sheet.getLastColumn()).fill('');
    setRowValue_(row, map, 'ROLE', role.role);
    setRowValue_(row, map, 'DISPLAY NAME', role.displayName);
    setRowValue_(row, map, 'PERMISSIONS', role.permissions.join(', '));
    setRowValue_(row, map, 'GLOBAL SCOPE', role.globalScope);
    setRowValue_(row, map, 'ACTIVE', true);
    setRowValue_(row, map, 'SYSTEM ROLE', true);
    setRowValue_(row, map, 'CREATED AT', now);
    setRowValue_(row, map, 'CREATED BY', owner);
    setRowValue_(row, map, 'UPDATED AT', now);
    setRowValue_(row, map, 'UPDATED BY', owner);
    return row;
  });
  if (rows.length) {
    ensureSheetCapacity_(sheet, rows.length);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, sheet.getLastColumn()).setValues(rows);
  }
  existing.forEach((row, index) => {
    const rawRole = normalizeLookupText_(getRowValue_(row, map, 'ROLE')).replace(/[\s-]+/g, '_');
    const role = canonicalRoleCode_(rawRole);
    if (!role) return;
    let changed = false;
    // Migrasi aman: APPROVER_OH -> OH. Permission dan scope tidak berubah.
    if (rawRole !== role && map['ROLE']) {
      sheet.getRange(index + 2, map['ROLE']).setValue(role);
      changed = true;
    }
    if (role === 'OH' && boolValue_(getRowValue_(row, map, 'SYSTEM ROLE')) && map['DISPLAY NAME'] &&
        cleanText_(getRowValue_(row, map, 'DISPLAY NAME')) !== 'OH') {
      sheet.getRange(index + 2, map['DISPLAY NAME']).setValue('OH');
      changed = true;
    }
    if (map['GLOBAL SCOPE'] && boolValue_(getRowValue_(row, map, 'GLOBAL SCOPE'))) {
      sheet.getRange(index + 2, map['GLOBAL SCOPE']).setValue(false);
      changed = true;
    }
    const permissions = splitAccessValues_(getRowValue_(row, map, 'PERMISSIONS'))
      .map(normalizeLookupText_).filter(permission => Object.prototype.hasOwnProperty.call(KM_RBAC.PERMISSIONS, permission));
    const normalizedPermissions = [...new Set(permissions)].join(', ');
    if (map['PERMISSIONS'] && cleanText_(getRowValue_(row, map, 'PERMISSIONS')) !== normalizedPermissions) {
      sheet.getRange(index + 2, map['PERMISSIONS']).setValue(normalizedPermissions);
      changed = true;
    }
    // Existing role tetap memakai permission tersimpan. Permission bawaan hanya
    // dipakai untuk row baru agar setup/maintenance tidak membatalkan edit admin.
    if (changed) {
      if (map['UPDATED AT']) sheet.getRange(index + 2, map['UPDATED AT']).setValue(now);
      if (map['UPDATED BY']) sheet.getRange(index + 2, map['UPDATED BY']).setValue(owner);
    }
  });
  putCachedJson_('roles', [], 1);
  KM_REQUEST_ROLE_DIRECTORY_ = null;
  KM_REQUEST_ACTIVE_ROLE_NAMES_ = null;
  return sheet;
}

function setupRbacDatabase_(ss) {
  return setupAuthDatabase_(ss).users;
}


function initializeEnterpriseAuth_() {
  const activeEmail = getSheetUserEmail_();
  const ownerEmail = getSpreadsheetOwnerEmail_();
  if (!activeEmail || (ownerEmail && activeEmail !== ownerEmail)) {
    throw new Error('Setup autentikasi hanya boleh dijalankan langsung oleh owner Spreadsheet dari Apps Script editor.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    getOrCreateAuthPepper_();
    const { users } = setupAuthDatabase_(ss);
    const map = getHeaderMap_(users);
    const values = users.getLastRow() < 2 ? [] : users.getRange(2, 1, users.getLastRow() - 1, users.getLastColumn()).getValues();
    let index = values.findIndex(row => normalizeEmail_(getRowValue_(row, map, 'EMAIL')) === activeEmail);
    const ownerRow = index >= 0 ? values[index] : null;
    const ownerAuthReady = Boolean(ownerRow &&
      isActiveAccessValue_(getRowValue_(ownerRow, map, 'ACTIVE')) &&
      normalizeRoles_(getRowValue_(ownerRow, map, 'ROLES')).includes('SUPER_ADMIN') &&
      cleanText_(getRowValue_(ownerRow, map, 'USER ID')) &&
      normalizeUsername_(getRowValue_(ownerRow, map, 'USERNAME')) &&
      cleanText_(getRowValue_(ownerRow, map, 'PASSWORD HASH')) &&
      cleanText_(getRowValue_(ownerRow, map, 'PASSWORD SALT'))
    );

    if (ownerAuthReady) {
      return {
        success: true,
        initialized: true,
        email: activeEmail,
        username: normalizeUsername_(getRowValue_(ownerRow, map, 'USERNAME')),
        temporaryPassword: '',
        mustChangePassword: boolValue_(getRowValue_(ownerRow, map, 'MUST CHANGE PASSWORD')),
        message: 'Enterprise Auth owner sudah aktif. Tidak ada password yang diubah.'
      };
    }

    const now = new Date();
    const temporaryPassword = generateTemporaryPassword_();
    const credential = createTemporaryPasswordCredential_(temporaryPassword);
    const row = index >= 0 ? values[index].slice() : new Array(users.getLastColumn()).fill('');
    const username = uniqueUsernameCandidate_(activeEmail.split('@')[0] || 'admin', values, map, index);
    setRowValue_(row, map, 'EMAIL', activeEmail);
    setRowValue_(row, map, 'ROLES', 'SUPER_ADMIN');
    setRowValue_(row, map, 'OPERATING POINTS', '*');
    setRowValue_(row, map, 'ACTIVE', true);
    setRowValue_(row, map, 'DISPLAY NAME', cleanText_(getRowValue_(row, map, 'DISPLAY NAME')) || 'Super Admin');
    setRowValue_(row, map, 'CREATED AT', getRowValue_(row, map, 'CREATED AT') || now);
    setRowValue_(row, map, 'CREATED BY', getRowValue_(row, map, 'CREATED BY') || activeEmail);
    setRowValue_(row, map, 'UPDATED AT', now);
    setRowValue_(row, map, 'UPDATED BY', activeEmail);
    setRowValue_(row, map, 'USER ID', cleanText_(getRowValue_(row, map, 'USER ID')) || newUserId_());
    setRowValue_(row, map, 'USERNAME', username);
    setRowValue_(row, map, 'PASSWORD HASH', credential.hash);
    setRowValue_(row, map, 'PASSWORD SALT', credential.salt);
    setRowValue_(row, map, 'PASSWORD ITERATIONS', credential.iterations);
    setRowValue_(row, map, 'PASSWORD ALGORITHM', credential.algorithm);
    setRowValue_(row, map, 'PASSWORD CHANGED AT', now);
    setRowValue_(row, map, 'MUST CHANGE PASSWORD', true);
    setRowValue_(row, map, 'FAILED LOGIN COUNT', 0);
    setRowValue_(row, map, 'LOCKED UNTIL', '');
    setRowValue_(row, map, 'SESSION VERSION', Math.max(1, Number(getRowValue_(row, map, 'SESSION VERSION')) || 1));

    const targetRow = index >= 0 ? index + 2 : users.getLastRow() + 1;
    ensureSheetCapacity_(users, index >= 0 ? 0 : 1);
    users.getRange(targetRow, 1, 1, users.getLastColumn()).setValues([row]);
    SpreadsheetApp.flush();
    invalidateSecurityCaches_();
    writeAuthLogUnsafe_({ identifier: activeEmail, userId: getRowValue_(row, map, 'USER ID'), email: activeEmail,
      action: 'AUTH_INITIALIZED', status: 'SUCCESS', note: 'Initial SUPER_ADMIN credential dibuat.' });
    runNonCritical_('hapus proteksi otomatis lama', removeAutomaticSheetProtections_);
    return {
      success: true,
      initialized: true,
      email: activeEmail,
      username,
      temporaryPassword,
      mustChangePassword: true,
      message: 'Simpan password sementara ini dengan aman. Password hanya ditampilkan pada hasil eksekusi ini.'
    };
  } finally {
    lock.releaseLock();
  }
}


function initializeKalibrasiKmApplication_() {
  const auth = initializeEnterpriseAuth_();
  const activeEmail = getSheetUserEmail_();
  const ownerEmail = getSpreadsheetOwnerEmail_();
  if (!activeEmail || (ownerEmail && activeEmail !== ownerEmail)) {
    throw new Error('Setup aplikasi hanya boleh dijalankan langsung oleh owner Spreadsheet dari Apps Script editor.');
  }

  const ownerUser = readUserDirectory_(true).find(user =>
    user.email === activeEmail && user.active && user.roles.includes('SUPER_ADMIN')
  );
  if (!ownerUser) {
    throw new Error('Setup auth belum menghasilkan akun SUPER_ADMIN owner yang aktif.');
  }

  const actor = buildActorFromUser_(ownerUser, 'SHEET');
  KM_REQUEST_ACTOR_ = actor;
  const database = setupDatabaseKalibrasiKMCore_(actor);
  const inspection = inspectDatabaseKalibrasiKm_();
  return {
    success: true,
    auth: auth,
    database: database,
    inspection: inspection,
    nextStep: auth.temporaryPassword
      ? 'Login menggunakan email/username owner dan password sementara, lalu ganti password.'
      : 'Login menggunakan credential owner yang sudah aktif.'
  };
}


function setupAplikasiKalibrasiKM_() {
  return initializeKalibrasiKmApplication_();
}


function periksaDatabaseKalibrasiKM_() {
  const actor = requirePermission_('ADMIN_SETUP', '', { source: 'SHEET' });
  KM_REQUEST_ACTOR_ = actor;
  return inspectDatabaseKalibrasiKm_();
}

function inspectCanonicalKeyRows_(sheet, keyHeader, knownKeyColumn) {
  if (!sheet || sheet.getLastRow() < 2) {
    return { keyHeader: keyHeader, recordCount: 0, gapRowCount: 0, orphanRowCount: 0, duplicateKeyCount: 0, gapRows: [], duplicateKeyRows: [] };
  }
  const keyColumn = Number(knownKeyColumn || 0) || getHeaderMap_(sheet)[keyHeader];
  if (!keyColumn) {
    return { keyHeader: keyHeader, recordCount: 0, gapRowCount: 0, orphanRowCount: 0, duplicateKeyCount: 0, gapRows: [], duplicateKeyRows: [] };
  }
  const rowCount = sheet.getLastRow() - 1;
  // Fast path untuk sheet sehat: baca kolom key saja. Seluruh row baru dibaca
  // bila memang ada key kosong dan perlu dibedakan antara gap vs orphan.
  const keyValues = sheet.getRange(2, keyColumn, rowCount, 1).getDisplayValues();
  const gaps = [];
  const duplicateRows = [];
  const seenKeys = new Set();
  let records = 0;
  keyValues.forEach((row, offset) => {
    const key = cleanText_(row[0]);
    if (key) {
      records++;
      if (seenKeys.has(key)) duplicateRows.push(offset + 2);
      else seenKeys.add(key);
      return;
    }
    gaps.push(offset + 2);
  });
  let orphans = 0;
  if (gaps.length) {
    const values = sheet.getRange(2, 1, rowCount, sheet.getLastColumn()).getDisplayValues();
    gaps.forEach(rowNumber => {
      if ((values[rowNumber - 2] || []).some(value => cleanText_(value))) orphans++;
    });
  }
  return {
    keyHeader: keyHeader,
    recordCount: records,
    gapRowCount: gaps.length,
    orphanRowCount: orphans,
    duplicateKeyCount: duplicateRows.length,
    gapRows: gaps.slice(0, 100),
    duplicateKeyRows: duplicateRows.slice(0, 100)
  };
}

function inspectDatabaseStructure_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const specs = enterpriseSheetSchemaSpecs_().map(spec => ({ name: spec[0], headers: spec[1] }));
  const sheets = specs.map(spec => {
    const sheet = ss.getSheetByName(spec.name);
    if (!sheet) {
      return { name: spec.name, exists: false, rowCount: 0, missingHeaders: spec.headers.slice(), duplicateHeaders: [], headerOrderReady: false, extraHeaders: [] };
    }
    const headers = sheet.getLastColumn()
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(cleanText_)
      : [];
    const normalized = headers.map(normalizeHeaderKey_).filter(Boolean);
    const missingHeaders = spec.headers.filter(header => !normalized.includes(normalizeHeaderKey_(header)));
    const duplicateHeaders = [...new Set(normalized.filter((header, index) => normalized.indexOf(header) !== index))];
    const expectedPositions = spec.headers.map(header => normalized.indexOf(normalizeHeaderKey_(header)));
    const headerOrderReady = expectedPositions.every((position, index) =>
      position >= 0 && (index === 0 || position > expectedPositions[index - 1])
    );
    const expectedKeys = new Set(spec.headers.map(normalizeHeaderKey_));
    const extraHeaders = headers.filter(header => cleanText_(header) && !expectedKeys.has(normalizeHeaderKey_(header)));
    const keyHeader = spec.name === KM_APP.SHEETS.SESSIONS ? 'SESSION ID'
      : spec.name === KM_APP.SHEETS.AUTH_LOG ? 'EVENT ID' : '';
    const keyColumn = keyHeader
      ? headers.findIndex(header => normalizeHeaderKey_(header) === normalizeHeaderKey_(keyHeader)) + 1
      : 0;
    const keyQuality = keyHeader ? inspectCanonicalKeyRows_(sheet, keyHeader, keyColumn) : {};
    return Object.assign({
      name: spec.name,
      exists: true,
      rowCount: Math.max(0, sheet.getLastRow() - 1),
      missingHeaders,
      duplicateHeaders,
      headerOrderReady,
      extraHeaders
    }, keyQuality);
  });

  const inventory = auditDatabaseSheetInventory_(ss);
  return {
    ready: sheets.every(item => item.exists && !item.missingHeaders.length &&
      !item.duplicateHeaders.length && item.headerOrderReady && !Number(item.gapRowCount || 0) &&
      !Number(item.duplicateKeyCount || 0)),
    spreadsheetName: ss.getName(),
    timezone: Session.getScriptTimeZone(),
    authPepperConfigured: Boolean(PropertiesService.getScriptProperties().getProperty(KM_AUTH.PEPPER_PROPERTY)),
    schemaVersion: KM_DATABASE_SCHEMA_VERSION,
    sheets: sheets,
    inventory: inventory
  };
}

function inspectDatabaseKalibrasiKm_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const structure = inspectDatabaseStructure_(ss);
  const sheets = structure.sheets;

  // MASTER_UNIT and snapshot audits use the same trusted ledgers. Build them
  // once for this request so large history sheets are not read twice.
  const sharedStatus3LLog = ss.getSheetByName(UNIT_STATUS_3L_LOG_SHEET);
  const sharedTrustedStatus = buildTrustedStatus3LStateIndex_(sharedStatus3LLog);
  const sharedTrustedFinals = buildTrustedFinalSnapshotIndex_(sharedTrustedStatus);
  const sharedAuditContext = {
    trustedStatus: sharedTrustedStatus,
    trustedFinals: sharedTrustedFinals
  };
  const masterUnit = auditMasterUnitKmDatabase_(sharedAuditContext);
  const historyContext = auditCalibrationHistoryContextIntegrity_();
  const repairSnapshotIntegrity = auditRepairSnapshotIntegrity_(sharedAuditContext);
  const analyticsLedger = auditAnalyticsCalibrationLedger_();
  const ready = sheets.every(item => item.exists && !item.missingHeaders.length && !item.duplicateHeaders.length &&
      item.headerOrderReady && !Number(item.gapRowCount || 0) && !Number(item.duplicateKeyCount || 0)) &&
    !masterUnit.missingColumns.length && !masterUnit.duplicateHeaders.length &&
    !masterUnit.invalidNopolFormatRows.length &&
    masterUnit.workflowColumnsReady && masterUnit.statusValidationReady &&
    masterUnit.historySheetReady && masterUnit.historyHidden &&
    masterUnit.statusHistorySheetReady && masterUnit.status3LHistorySheetReady &&
    masterUnit.editTriggerReady &&
    Boolean(historyContext.ready) && Boolean(repairSnapshotIntegrity.ready) && Boolean(analyticsLedger.ready) &&
    !(masterUnit.status3LAttributionMismatchRows || []).length &&
    !(masterUnit.status3LUnsupportedDoneRows || []).length &&
    !(masterUnit.status3LStateMismatchRows || []).length &&
    !(masterUnit.status3LStateUnresolvedRows || []).length;
  const report = {
    ready,
    spreadsheetName: structure.spreadsheetName,
    timezone: structure.timezone,
    authPepperConfigured: structure.authPepperConfigured,
    schemaVersion: structure.schemaVersion,
    sheets,
    inventory: structure.inventory,
    masterUnit,
    historyContext,
    repairSnapshotIntegrity,
    analyticsLedger,
    recommendation: ready
      ? 'Struktur database, workflow evidence-first, RBAC backend, dan clean analytics ledger siap digunakan.'
      : 'Jalankan Administration > Database Health > Full Safe Repair, lalu audit kembali sampai seluruh critical issue bernilai 0.'
  };
  console.log(JSON.stringify({
    audit: 'DATABASE_HEALTH', ready: report.ready,
    sheets: report.sheets.length,
    historyReady: Boolean(report.historyContext && report.historyContext.ready),
    snapshotReady: Boolean(report.repairSnapshotIntegrity && report.repairSnapshotIntegrity.ready),
    analyticsReady: Boolean(report.analyticsLedger && report.analyticsLedger.ready)
  }));
  return report;
}

function resetOwnerEnterprisePassword_() {
  const activeEmail = getSheetUserEmail_();
  const ownerEmail = getSpreadsheetOwnerEmail_();
  if (!activeEmail || !ownerEmail || activeEmail !== ownerEmail) {
    throw new Error('Recovery hanya boleh dijalankan langsung oleh owner Spreadsheet dari Apps Script editor.');
  }
  getAuthPepper_();
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.USERS);
    if (!sheet || sheet.getLastRow() < 2) throw new Error('KM_USERS belum tersedia. Jalankan initializeEnterpriseAuth_().');
    ensureHeaders_(sheet, KM_APP.USER_HEADERS);
    const map = getHeaderMap_(sheet);
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const index = values.findIndex(row => normalizeEmail_(getRowValue_(row, map, 'EMAIL')) === ownerEmail);
    if (index < 0) throw new Error('Akun owner belum ditemukan di KM_USERS.');
    const row = values[index].slice();
    const temporaryPassword = generateTemporaryPassword_();
    const credential = createTemporaryPasswordCredential_(temporaryPassword);
    const now = new Date();
    const userId = cleanText_(getRowValue_(row, map, 'USER ID')) || newUserId_();
    const nextVersion = Math.max(1, Number(getRowValue_(row, map, 'SESSION VERSION')) || 1) + 1;
    setRowValue_(row, map, 'USER ID', userId);
    setRowValue_(row, map, 'ROLES', normalizeRoles_(getRowValue_(row, map, 'ROLES')).includes('SUPER_ADMIN')
      ? normalizeRoles_(getRowValue_(row, map, 'ROLES')).join(', ') : 'SUPER_ADMIN');
    setRowValue_(row, map, 'OPERATING POINTS', '*');
    setRowValue_(row, map, 'ACTIVE', true);
    setRowValue_(row, map, 'PASSWORD HASH', credential.hash);
    setRowValue_(row, map, 'PASSWORD SALT', credential.salt);
    setRowValue_(row, map, 'PASSWORD ITERATIONS', credential.iterations);
    setRowValue_(row, map, 'PASSWORD ALGORITHM', credential.algorithm);
    setRowValue_(row, map, 'PASSWORD CHANGED AT', now);
    setRowValue_(row, map, 'MUST CHANGE PASSWORD', true);
    setRowValue_(row, map, 'FAILED LOGIN COUNT', 0);
    setRowValue_(row, map, 'LOCKED UNTIL', '');
    setRowValue_(row, map, 'SESSION VERSION', nextVersion);
    setRowValue_(row, map, 'UPDATED AT', now);
    setRowValue_(row, map, 'UPDATED BY', ownerEmail);
    sheet.getRange(index + 2, 1, 1, sheet.getLastColumn()).setValues([row]);
    revokeAllUserSessions_(userId, 'OWNER_BREAK_GLASS_RESET');
    invalidateUserDirectoryCache_();
    writeAuthLogUnsafe_({ identifier: ownerEmail, userId, email: ownerEmail, action: 'OWNER_BREAK_GLASS_RESET', status: 'SUCCESS' });
    const result = {
      success: true,
      email: ownerEmail,
      username: normalizeUsername_(getRowValue_(row, map, 'USERNAME')),
      temporaryPassword,
      mustChangePassword: true
    };
    console.log(
      'PASSWORD SEMENTARA ADMIN (RAHASIA)\n' +
      'Email: ' + result.email + '\n' +
      'Username: ' + result.username + '\n' +
      'Password sementara: ' + result.temporaryPassword + '\n' +
      'Password wajib diganti setelah login pertama.'
    );
    return result;
  } finally {
    lock.releaseLock();
  }
}


function buatUlangPasswordAdmin_() {
  return resetOwnerEnterprisePassword_();
}

function isAutomaticSheetProtectionDescription_(description) {
  const value = String(description || '');
  return [
    'KM_AUTH:',
    'KM_HISTORY:IMMUTABLE:',
    'KM_ANALYTICS:APP_MANAGED:',
    'KM_RBAC:STATUS_3L',
    'KM_RBAC:STATUS_UNIT',
    'KM_MASTER_UNIT_ADMIN:IMMUTABLE:',
    'KM_MASTER_UNIT_ADMIN:METADATA:'
  ].some(prefix => value.indexOf(prefix) === 0);
}

/** Remove only protections previously created by this application. */
function removeAutomaticSheetProtections_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let removedSheetProtections = 0;
  let removedRangeProtections = 0;
  const failures = [];
  ss.getSheets().forEach(sheet => {
    [
      { type: SpreadsheetApp.ProtectionType.SHEET, kind: 'sheet' },
      { type: SpreadsheetApp.ProtectionType.RANGE, kind: 'range' }
    ].forEach(entry => {
      sheet.getProtections(entry.type)
        .filter(item => isAutomaticSheetProtectionDescription_(item.getDescription()))
        .forEach(item => {
          try {
            item.remove();
            if (entry.kind === 'sheet') removedSheetProtections++;
            else removedRangeProtections++;
          } catch (error) {
            failures.push({
              sheet: sheet.getName(),
              description: String(item.getDescription() || ''),
              message: String(error && error.message ? error.message : error)
            });
          }
        });
    });
  });
  return {
    success: failures.length === 0,
    mode: 'BACKEND_RBAC_ONLY',
    automaticProtectionCreation: false,
    removedSheetProtections,
    removedRangeProtections,
    failures
  };
}

function syncRbacSheetProtections_() {
  return Object.assign({ updated: 0 }, removeAutomaticSheetProtections_());
}

function getOrCreateAuthPepper_() {
  const props = PropertiesService.getScriptProperties();
  let pepper = props.getProperty(KM_AUTH.PEPPER_PROPERTY);
  if (!pepper) {
    pepper = randomBase64Url_(32);
    props.setProperty(KM_AUTH.PEPPER_PROPERTY, pepper);
  }
  return pepper;
}

function getAuthPepper_() {
  const pepper = PropertiesService.getScriptProperties().getProperty(KM_AUTH.PEPPER_PROPERTY);
  if (!pepper) throw new Error('AUTH_NOT_INITIALIZED: jalankan initializeKalibrasiKmApplication_() dari Apps Script editor.');
  return pepper;
}

function randomBase64Url_(bytesLength) {
  let seed = '';
  while (seed.length < bytesLength * 4) seed += Utilities.getUuid() + ':' + new Date().getTime() + ':' + Math.random();
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed, Utilities.Charset.UTF_8);
  let bytes = digest;
  while (bytes.length < bytesLength) {
    bytes = bytes.concat(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
      bytesToHex_(bytes) + Utilities.getUuid(), Utilities.Charset.UTF_8));
  }
  return Utilities.base64EncodeWebSafe(bytes.slice(0, bytesLength)).replace(/=+$/g, '');
}

function bytesToHex_(bytes) {
  return bytes.map(byte => ('0' + ((byte < 0 ? byte + 256 : byte) & 255).toString(16)).slice(-2)).join('');
}

function sha256Hex_(value) {
  return bytesToHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    String(value == null ? '' : value), Utilities.Charset.UTF_8));
}

function constantTimeEquals_(a, b) {
  a = String(a || ''); b = String(b || '');
  let diff = a.length ^ b.length;
  const len = Math.max(a.length, b.length);
  for (let i = 0; i < len; i++) diff |= (a.charCodeAt(i % Math.max(a.length, 1)) || 0) ^ (b.charCodeAt(i % Math.max(b.length, 1)) || 0);
  return diff === 0;
}

function base64WebSafeDecode_(value) {
  let text = String(value || '');
  while (text.length % 4) text += '=';
  return Utilities.base64DecodeWebSafe(text);
}

function pbkdf2Sha256Hex_(password, saltBase64, iterations) {
  const pepper = getAuthPepper_();
  const keyBytes = Utilities.newBlob(String(password) + ':' + pepper).getBytes();
  const saltBytes = base64WebSafeDecode_(saltBase64);
  const block = saltBytes.concat([0, 0, 0, 1]);
  let u = Utilities.computeHmacSha256Signature(block, keyBytes);
  const t = u.map(byte => (byte < 0 ? byte + 256 : byte) & 255);
  const rounds = Math.max(1, Number(iterations) || KM_AUTH.PASSWORD_ITERATIONS);
  for (let i = 1; i < rounds; i++) {
    u = Utilities.computeHmacSha256Signature(u, keyBytes);
    for (let j = 0; j < t.length; j++) t[j] ^= ((u[j] < 0 ? u[j] + 256 : u[j]) & 255);
  }
  return t.map(byte => ('0' + (byte & 255).toString(16)).slice(-2)).join('');
}

function createPasswordCredential_(password) {
  const salt = randomBase64Url_(24);
  const iterations = KM_AUTH.PASSWORD_ITERATIONS;
  return {
    salt,
    iterations,
    algorithm: 'PBKDF2_SHA256_V1',
    hash: pbkdf2Sha256Hex_(password, salt, iterations)
  };
}

function temporaryPasswordHash_(password, saltBase64) {
  const pepper = getAuthPepper_();
  const valueBytes = Utilities.newBlob(String(password) + ':' + String(saltBase64 || '')).getBytes();
  const keyBytes = Utilities.newBlob(pepper).getBytes();
  return bytesToHex_(Utilities.computeHmacSha256Signature(valueBytes, keyBytes));
}

function createTemporaryPasswordCredential_(password) {
  const salt = randomBase64Url_(24);
  return {
    salt,
    iterations: 1,
    algorithm: 'TEMP_HMAC_SHA256_V1',
    hash: temporaryPasswordHash_(password, salt)
  };
}

function verifyPassword_(password, user) {
  if (!user || !user.passwordHash || !user.passwordSalt) return false;
  const algorithm = cleanText_(user.passwordAlgorithm || 'PBKDF2_SHA256_V1').toUpperCase();
  const candidate = algorithm === 'TEMP_HMAC_SHA256_V1'
    ? temporaryPasswordHash_(password, user.passwordSalt)
    : pbkdf2Sha256Hex_(password, user.passwordSalt, user.passwordIterations);
  return constantTimeEquals_(candidate, user.passwordHash);
}

function passwordPolicyErrors_(password, user) {
  const value = String(password || '');
  const errors = [];
  if (value.length < KM_AUTH.PASSWORD_MIN_LENGTH) errors.push('minimal ' + KM_AUTH.PASSWORD_MIN_LENGTH + ' karakter');
  if (value.length > KM_AUTH.PASSWORD_MAX_LENGTH) errors.push('maksimal ' + KM_AUTH.PASSWORD_MAX_LENGTH + ' karakter');
  if (!/[a-z]/.test(value)) errors.push('huruf kecil');
  if (!/[A-Z]/.test(value)) errors.push('huruf besar');
  if (!/\d/.test(value)) errors.push('angka');
  const lowered = value.toLowerCase();
  [user && user.username, user && user.email && user.email.split('@')[0]].filter(item => item && item.length >= 4)
    .forEach(item => { if (lowered.indexOf(String(item).toLowerCase()) >= 0) errors.push('tidak boleh mengandung username/email'); });
  return [...new Set(errors)];
}

function generateTemporaryPassword_() {
  const upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ';
  const lower = 'abcdefghijkmnopqrstuvwxyz';
  const digits = '23456789';
  const symbols = '!@#$%*-_=+';
  const all = upper + lower + digits + symbols;
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + ':' + new Date().getTime() + ':' + Math.random(), Utilities.Charset.UTF_8);
  const pick = (chars, i) => chars[((bytes[i % bytes.length] < 0 ? bytes[i % bytes.length] + 256 : bytes[i % bytes.length]) & 255) % chars.length];
  let chars = [pick(upper, 0), pick(lower, 1), pick(digits, 2), pick(symbols, 3)];
  for (let i = 4; i < KM_AUTH.TEMP_PASSWORD_LENGTH; i++) chars.push(pick(all, i));
  // Deterministic shuffle from random bytes.
  for (let i = chars.length - 1; i > 0; i--) {
    const b = (bytes[(i + 7) % bytes.length] < 0 ? bytes[(i + 7) % bytes.length] + 256 : bytes[(i + 7) % bytes.length]) & 255;
    const j = b % (i + 1); [chars[i], chars[j]] = [chars[j], chars[i]];
  }
  return chars.join('');
}

function newUserId_() {
  return 'USR-' + Utilities.getUuid();
}

function uniqueUsernameCandidate_(candidate, values, map, ignoreIndex) {
  const base = normalizeUsername_(candidate || 'user').replace(/[^a-z0-9._-]/g, '') || 'user';
  const existing = new Set(values.map((row, index) => index === ignoreIndex ? '' : normalizeUsername_(getRowValue_(row, map, 'USERNAME'))).filter(Boolean));
  if (!existing.has(base)) return base;
  for (let i = 2; i < 10000; i++) if (!existing.has(base + i)) return base + i;
  throw new Error('Tidak dapat membuat username unik.');
}

function sessionCacheKey_(sessionId) {
  return KM_AUTH.SESSION_CACHE_PREFIX + sessionId;
}

function sessionRevokedCacheKey_(sessionId) {
  return KM_AUTH.SESSION_REVOKED_CACHE_PREFIX + sessionId;
}

function markSessionRevokedInCache_(sessionId) {
  if (!sessionId) return;
  const cache = CacheService.getScriptCache();
  cache.remove(sessionCacheKey_(sessionId));
  // Tombstone mencegah request paralel yang sudah membaca sesi ACTIVE
  // menghidupkan kembali cache sesudah logout/revoke selesai.
  cache.put(
    sessionRevokedCacheKey_(sessionId),
    '1',
    Math.max(KM_AUTH.SESSION_CACHE_SECONDS + 60, 600)
  );
}

function parseSessionToken_(token) {
  const parts = String(token || '').split('.');
  if (parts.length !== 2 || !/^SES-[A-Za-z0-9-]+$/.test(parts[0]) || parts[1].length < 20) return null;
  return { sessionId: parts[0], secret: parts[1] };
}

function createSessionForUser_(user, clientInfo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, KM_APP.SHEETS.SESSIONS);
  ensureHeaders_(sheet, KM_AUTH.SESSION_HEADERS);
  const map = getHeaderMap_(sheet);
  const now = new Date();
  const expiresAt = new Date(now.getTime() + KM_AUTH.SESSION_TTL_HOURS * 3600000);
  const sessionId = 'SES-' + Utilities.getUuid();
  const secret = randomBase64Url_(32);
  const tokenHash = sha256Hex_(secret);
  const userAgentHash = sha256Hex_(clientInfo && clientInfo.userAgent || '').slice(0, 32);
  const row = new Array(sheet.getLastColumn()).fill('');
  setRowValue_(row, map, 'SESSION ID', sessionId);
  setRowValue_(row, map, 'USER ID', user.userId);
  setRowValue_(row, map, 'TOKEN HASH', tokenHash);
  setRowValue_(row, map, 'CREATED AT', now);
  setRowValue_(row, map, 'EXPIRES AT', expiresAt);
  setRowValue_(row, map, 'LAST ACTIVITY AT', now);
  setRowValue_(row, map, 'STATUS', 'ACTIVE');
  setRowValue_(row, map, 'SESSION VERSION', user.sessionVersion);
  setRowValue_(row, map, 'USER AGENT HASH', userAgentHash);
  const rowNumber = appendCanonicalRowUnsafe_(sheet, row, 'SESSION ID');
  const cachePayload = {
    rowNumber, sessionId, userId: user.userId, tokenHash, createdAt: now.toISOString(),
    expiresAt: expiresAt.toISOString(), lastActivityAt: now.toISOString(), status: 'ACTIVE',
    sessionVersion: user.sessionVersion, userAgentHash, lastSheetTouchAt: now.toISOString()
  };
  const cache = CacheService.getScriptCache();
  cache.remove(sessionRevokedCacheKey_(sessionId));
  cache.put(sessionCacheKey_(sessionId), JSON.stringify(cachePayload), KM_AUTH.SESSION_CACHE_SECONDS);
  return { token: sessionId + '.' + secret, sessionId, expiresAt: expiresAt.toISOString() };
}

function findUniqueRowByColumnValue_(sheet, column, value) {
  if (!sheet || !column || !cleanText_(value) || sheet.getLastRow() < 2) return 0;
  const matches = sheet.getRange(2, column, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(value))
    .matchEntireCell(true)
    .matchCase(true)
    .findAll();
  return matches.length === 1 ? matches[0].getRow() : 0;
}

function readSessionById_(sessionId) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(sessionCacheKey_(sessionId));
  if (cached) {
    try { return JSON.parse(cached); } catch (error) {}
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.SESSIONS);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const map = getHeaderMap_(sheet);
  const idCol = map['SESSION ID'];
  if (!idCol) return null;
  const rowNumber = findUniqueRowByColumnValue_(sheet, idCol, sessionId);
  if (!rowNumber) return null;
  const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  const result = {
    rowNumber,
    sessionId,
    userId: cleanText_(getRowValue_(row, map, 'USER ID')),
    tokenHash: cleanText_(getRowValue_(row, map, 'TOKEN HASH')),
    createdAt: dateToIso_(getRowValue_(row, map, 'CREATED AT')),
    expiresAt: dateToIso_(getRowValue_(row, map, 'EXPIRES AT')),
    lastActivityAt: dateToIso_(getRowValue_(row, map, 'LAST ACTIVITY AT')),
    status: normalizeLookupText_(getRowValue_(row, map, 'STATUS')),
    sessionVersion: Number(getRowValue_(row, map, 'SESSION VERSION')) || 1,
    userAgentHash: cleanText_(getRowValue_(row, map, 'USER AGENT HASH')),
    lastSheetTouchAt: dateToIso_(getRowValue_(row, map, 'LAST ACTIVITY AT'))
  };
  cache.put(sessionCacheKey_(sessionId), JSON.stringify(result), KM_AUTH.SESSION_CACHE_SECONDS);
  return result;
}

function resolveSessionRowNumber_(sheet, map, session) {
  if (!sheet || !session || !session.sessionId || !map['SESSION ID']) return 0;
  const cachedRow = Number(session.rowNumber) || 0;
  if (cachedRow >= 2 && cachedRow <= sheet.getMaxRows()) {
    const cachedId = cleanText_(sheet.getRange(cachedRow, map['SESSION ID']).getDisplayValue());
    if (cachedId === session.sessionId) return cachedRow;
  }
  const resolvedRow = findUniqueRowByColumnValue_(sheet, map['SESSION ID'], session.sessionId);
  if (resolvedRow) session.rowNumber = resolvedRow;
  return resolvedRow;
}

function dateToIso_(value) {
  const date = value instanceof Date ? value : parseDateTimeInput_(value);
  return date instanceof Date && !isNaN(date.getTime()) ? date.toISOString() : '';
}

function readSessionUserState_(userId) {
  const cached = KM_REQUEST_USER_DIRECTORY_ || getCachedJson_('users');
  if (Array.isArray(cached)) {
    const user = cached.find(item => item.userId === userId);
    if (user) return user;
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.USERS);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const map = getHeaderMap_(sheet);
  const idColumn = map['USER ID'];
  if (!idColumn) return null;
  const rowNumber = findUniqueRowByColumnValue_(sheet, idColumn, userId);
  if (!rowNumber) return null;
  const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  return {
    userId: cleanText_(getRowValue_(row, map, 'USER ID')),
    email: normalizeEmail_(getRowValue_(row, map, 'EMAIL')),
    active: isActiveAccessValue_(getRowValue_(row, map, 'ACTIVE')),
    sessionVersion: Math.max(1, Number(getRowValue_(row, map, 'SESSION VERSION')) || 1)
  };
}

function validateSessionToken_(token, options) {
  const parsed = parseSessionToken_(token);
  if (!parsed) throw new Error('SESSION_EXPIRED: sesi tidak valid. Silakan login kembali.');
  if (CacheService.getScriptCache().get(sessionRevokedCacheKey_(parsed.sessionId))) {
    throw new Error('SESSION_REVOKED: sesi sudah dicabut. Silakan login kembali.');
  }
  const session = readSessionById_(parsed.sessionId);
  if (!session || session.status !== 'ACTIVE' || !constantTimeEquals_(sha256Hex_(parsed.secret), session.tokenHash)) {
    throw new Error('SESSION_EXPIRED: sesi tidak valid atau sudah dicabut. Silakan login kembali.');
  }

  const now = new Date();
  const expiresAt = new Date(session.expiresAt);
  const lastActivity = new Date(session.lastActivityAt);
  const idleMs = KM_AUTH.IDLE_TIMEOUT_MINUTES * 60000;
  if (isNaN(expiresAt.getTime()) || now > expiresAt || isNaN(lastActivity.getTime()) || now.getTime() - lastActivity.getTime() >= idleMs) {
    revokeSessionByRecord_(session, 'EXPIRED');
    writeAuthLog_({ userId: session.userId, action: 'SESSION_EXPIRED', status: 'SUCCESS', sessionId: session.sessionId });
    throw new Error('SESSION_EXPIRED: sesi berakhir karena timeout. Silakan login kembali.');
  }

  const user = options && options.healthCheck
    ? readSessionUserState_(session.userId)
    : readUserDirectory_().find(item => item.userId === session.userId);
  if (!user || !user.active || user.sessionVersion !== session.sessionVersion) {
    revokeSessionByRecord_(session, 'USER_REVOKED');
    throw new Error('SESSION_REVOKED: akses akun berubah. Silakan login kembali.');
  }

  if (!(options && options.touchSession === false)) touchSession_(session, now);
  // Heartbeat hanya memerlukan validitas user/session. Hindari normalisasi role
  // dan penyusunan permission penuh pada request health yang rutin.
  if (options && options.healthCheck) {
    return {
      authorized: true,
      userId: user.userId,
      email: user.email,
      sessionId: session.sessionId,
      sessionExpiresAt: session.expiresAt
    };
  }
  const actor = buildActorFromUser_(user, 'WEB_APP', { sessionId: session.sessionId, expiresAt: session.expiresAt });
  return actor;
}

function touchSession_(session, now) {
  const cache = CacheService.getScriptCache();
  if (cache.get(sessionRevokedCacheKey_(session.sessionId))) return;
  const previousActivityAt = session.lastActivityAt;
  session.lastActivityAt = now.toISOString();
  const lastSheetTouch = new Date(session.lastSheetTouchAt || previousActivityAt);
  if (isNaN(lastSheetTouch.getTime()) || now.getTime() - lastSheetTouch.getTime() >= KM_AUTH.SESSION_TOUCH_MINUTES * 60000) {
    const lock = LockService.getScriptLock();
    let locked = false;
    try {
      locked = lock.tryLock(1500);
    } catch (error) {
      console.warn('Lock LAST ACTIVITY AT gagal: ' + String(error && error.message ? error.message : error));
    }
    if (locked) {
      try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.SESSIONS);
        if (sheet) {
          const map = getHeaderMap_(sheet);
          const rowNumber = resolveSessionRowNumber_(sheet, map, session);
          if (rowNumber && map['LAST ACTIVITY AT']) {
            sheet.getRange(rowNumber, map['LAST ACTIVITY AT']).setValue(now);
            session.lastSheetTouchAt = now.toISOString();
          }
        }
      } finally {
        lock.releaseLock();
      }
    } else {
      console.warn('Update LAST ACTIVITY AT ditunda karena transaksi auth lain sedang berjalan.');
    }
  }
  cache.put(sessionCacheKey_(session.sessionId), JSON.stringify(session), KM_AUTH.SESSION_CACHE_SECONDS);
}

function revokeSessionByRecord_(session, reason) {
  if (session && session.sessionId) markSessionRevokedInCache_(session.sessionId);
  const lock = acquireScriptLock_(7000, 'pencabutan sesi');
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.SESSIONS);
    if (sheet && session) {
      const map = getHeaderMap_(sheet);
      const rowNumber = resolveSessionRowNumber_(sheet, map, session);
      if (rowNumber) {
        const now = new Date();
        if (map['STATUS']) sheet.getRange(rowNumber, map['STATUS']).setValue('REVOKED');
        if (map['REVOKED AT']) sheet.getRange(rowNumber, map['REVOKED AT']).setValue(now);
        if (map['REVOKE REASON']) sheet.getRange(rowNumber, map['REVOKE REASON']).setValue(reason || 'REVOKED');
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function revokeAllUserSessions_(userId, reason) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.SESSIONS);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const map = getHeaderMap_(sheet);
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const values = range.getValues();
  const now = new Date();
  const revokedSessionIds = [];
  let count = 0;

  values.forEach(row => {
    if (cleanText_(getRowValue_(row, map, 'USER ID')) !== userId ||
        normalizeLookupText_(getRowValue_(row, map, 'STATUS')) !== 'ACTIVE') return;
    setRowValue_(row, map, 'STATUS', 'REVOKED');
    setRowValue_(row, map, 'REVOKED AT', now);
    setRowValue_(row, map, 'REVOKE REASON', reason || 'USER_SECURITY_CHANGE');
    const sessionId = cleanText_(getRowValue_(row, map, 'SESSION ID'));
    if (sessionId) revokedSessionIds.push(sessionId);
    count++;
  });

  if (count) {
    range.setValues(values);
    revokedSessionIds.forEach(sessionId => markSessionRevokedInCache_(sessionId));
  }
  return count;
}

function writeAuthLogUnsafe_(entry) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, KM_APP.SHEETS.AUTH_LOG);
  ensureHeaders_(sheet, KM_AUTH.AUTH_LOG_HEADERS);
  const map = getHeaderMap_(sheet);
  const row = new Array(sheet.getLastColumn()).fill('');
  setRowValue_(row, map, 'EVENT ID', 'AUTH-' + Utilities.getUuid());
  setRowValue_(row, map, 'TIMESTAMP', new Date());
  setRowValue_(row, map, 'IDENTIFIER', safeSheetText_(entry.identifier));
  setRowValue_(row, map, 'USER ID', safeSheetText_(entry.userId));
  setRowValue_(row, map, 'EMAIL', safeSheetText_(entry.email));
  setRowValue_(row, map, 'ACTION', safeSheetText_(entry.action));
  setRowValue_(row, map, 'STATUS', safeSheetText_(entry.status));
  setRowValue_(row, map, 'NOTE', safeSheetText_(entry.note));
  setRowValue_(row, map, 'USER AGENT HASH', safeSheetText_(entry.userAgentHash));
  setRowValue_(row, map, 'SESSION ID', safeSheetText_(entry.sessionId));
  return appendCanonicalRowUnsafe_(sheet, row, 'EVENT ID');
}

function writeAuthLog_(entry) {
  const lock = acquireScriptLock_(7000, 'penulisan auth log');
  try {
    return writeAuthLogUnsafe_(entry);
  } finally {
    lock.releaseLock();
  }
}

function unknownLoginThrottle_(identifier) {
  const key = KM_AUTH.UNKNOWN_LOGIN_PREFIX + sha256Hex_(identifier).slice(0, 32);
  const cache = CacheService.getScriptCache();
  const count = Number(cache.get(key)) || 0;
  cache.put(key, String(count + 1), KM_AUTH.UNKNOWN_LOGIN_CACHE_SECONDS);
  return count + 1;
}

function isUnknownLoginThrottled_(identifier) {
  const key = KM_AUTH.UNKNOWN_LOGIN_PREFIX + sha256Hex_(identifier).slice(0, 32);
  return (Number(CacheService.getScriptCache().get(key)) || 0) >= KM_AUTH.MAX_UNKNOWN_LOGIN_ATTEMPTS;
}


function findLoginUser_(normalizedIdentifier, forceRefresh) {
  const username = normalizeUsername_(normalizedIdentifier);
  let users = readUserDirectory_(Boolean(forceRefresh));
  let matches = users.filter(user =>
    user.email === normalizedIdentifier || user.username === username
  );

  if (!matches.length && !forceRefresh) {
    users = readUserDirectory_(true);
    matches = users.filter(user =>
      user.email === normalizedIdentifier || user.username === username
    );
  }

  return {
    users: users,
    user: matches.length === 1 ? matches[0] : null,
    ambiguous: matches.length > 1
  };
}

function loginCredentialVersion_(user) {
  return [
    user && user.userId,
    user && user.passwordHash,
    user && user.passwordSalt,
    user && user.passwordIterations,
    user && user.passwordAlgorithm
  ].map(value => String(value || '')).join('|');
}

function findUserRowById_(sheet, map, userId) {
  const column = map['USER ID'];
  const rowNumber = findUniqueRowByColumnValue_(sheet, column, userId);
  if (!rowNumber) return null;
  const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  return { rowNumber: rowNumber, row: row, user: userRecordFromRow_(row, map, rowNumber) };
}

function loginUser(identifier, password, clientInfo) {
  const rawIdentifier = String(identifier || '').trim();
  const normalizedIdentifier = normalizeEmail_(rawIdentifier);
  const rawPassword = String(password || '');
  const uaHash = sha256Hex_(String(clientInfo && clientInfo.userAgent || '').slice(0, 1024)).slice(0, 32);
  if (!normalizedIdentifier || !rawPassword) throw new Error('LOGIN_FAILED: email/username dan password wajib diisi.');
  if (rawIdentifier.length > 120 || rawPassword.length > KM_AUTH.PASSWORD_MAX_LENGTH) {
    throw new Error('LOGIN_FAILED: email/username atau password tidak valid.');
  }
  // Setelah beberapa percobaan untuk identifier yang tidak dikenal, hentikan
  // request sebelum pembacaan Sheet dan penulisan audit tambahan.
  if (isUnknownLoginThrottled_(normalizedIdentifier)) {
    throw new Error('LOGIN_RATE_LIMITED: Terlalu banyak percobaan untuk akun ini. Tunggu sekitar 15 menit lalu coba kembali.');
  }

  getAuthPepper_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let usersSheet = ss.getSheetByName(KM_APP.SHEETS.USERS);
  if (!usersSheet || usersSheet.getLastRow() < 2) {
    throw new Error('AUTH_NOT_INITIALIZED: jalankan initializeKalibrasiKmApplication_() dari Apps Script editor.');
  }
  let directory = findLoginUser_(normalizedIdentifier);
  let directoryUser = directory.user;
  if (!directoryUser || directory.ambiguous) {
    unknownLoginThrottle_(normalizedIdentifier);
    const auditLock = LockService.getScriptLock();
    if (auditLock.tryLock(1500)) {
      try {
        writeAuthLogUnsafe_({ identifier: normalizedIdentifier, action: 'LOGIN', status: 'FAILED', note: 'Invalid credential.', userAgentHash: uaHash });
      } finally {
        auditLock.releaseLock();
      }
    } else {
      console.warn('Audit login gagal tidak ditulis karena sistem sedang sibuk.');
    }
    throw new Error('LOGIN_FAILED: email/username atau password tidak valid.');
  }

  let map = getHeaderMap_(usersSheet);
  let rowNumber = directoryUser.rowNumber;
  let row = usersSheet.getRange(rowNumber, 1, 1, usersSheet.getLastColumn()).getValues()[0];
  let user = userRecordFromRow_(row, map, rowNumber);
  // Proteksi bila KM_USERS disortir ketika cache masih aktif.
  if (user.userId !== directoryUser.userId ||
      (user.email !== normalizedIdentifier && user.username !== normalizeUsername_(normalizedIdentifier))) {
    directory = findLoginUser_(normalizedIdentifier, true);
    directoryUser = directory.user;
    if (!directoryUser || directory.ambiguous) {
      throw new Error('LOGIN_FAILED: email/username atau password tidak valid.');
    }
    rowNumber = directoryUser.rowNumber;
    row = usersSheet.getRange(rowNumber, 1, 1, usersSheet.getLastColumn()).getValues()[0];
    user = userRecordFromRow_(row, map, rowNumber);
  }

  // PBKDF2 adalah operasi CPU paling mahal pada login. Jalankan sebelum lock
  // agar login satu pengguna tidak menghentikan request pengguna lain.
  const credentialVersion = loginCredentialVersion_(user);
  const passwordValid = Boolean(user.active) && verifyPassword_(rawPassword, user);

  const lock = acquireScriptLock_(7000, 'login');
  let authenticated = null;
  try {
    usersSheet = ss.getSheetByName(KM_APP.SHEETS.USERS);
    if (!usersSheet || usersSheet.getLastRow() < 2) {
      throw new Error('AUTH_NOT_INITIALIZED: jalankan initializeKalibrasiKmApplication_() dari Apps Script editor.');
    }
    map = getHeaderMap_(usersSheet);
    const fresh = findUserRowById_(usersSheet, map, directoryUser.userId);
    if (!fresh || loginCredentialVersion_(fresh.user) !== credentialVersion) {
      throw new Error('LOGIN_FAILED: data akun berubah. Silakan ulangi login.');
    }
    rowNumber = fresh.rowNumber;
    row = fresh.row;
    user = fresh.user;
    const now = new Date();
    const lockedUntil = user && user.lockedUntilRaw instanceof Date ? user.lockedUntilRaw : parseDateTimeInput_(user && user.lockedUntilRaw);
    if (!user || !user.active) {
      writeAuthLogUnsafe_({ identifier: normalizedIdentifier, userId: user && user.userId, email: user && user.email,
        action: 'LOGIN', status: 'FAILED', note: 'Inactive account.', userAgentHash: uaHash });
      throw new Error('LOGIN_FAILED: email/username atau password tidak valid.');
    }
    if (lockedUntil instanceof Date && lockedUntil.getTime() > now.getTime()) {
      const remainingMinutes = Math.max(1, Math.ceil((lockedUntil.getTime() - now.getTime()) / 60000));
      writeAuthLogUnsafe_({ identifier: normalizedIdentifier, userId: user.userId, email: user.email,
        action: 'LOGIN', status: 'LOCKED', note: 'Account temporarily locked.', userAgentHash: uaHash });
      throw new Error('LOGIN_LOCKED: Akun sementara dikunci. Coba kembali sekitar ' + remainingMinutes + ' menit lagi.');
    }

    if (!passwordValid) {
      const failedCount = (user.failedLoginCount || 0) + 1;
      setRowValue_(row, map, 'FAILED LOGIN COUNT', failedCount);
      if (failedCount >= KM_AUTH.MAX_FAILED_LOGINS) {
        setRowValue_(row, map, 'LOCKED UNTIL', new Date(now.getTime() + KM_AUTH.LOCK_MINUTES * 60000));
        setRowValue_(row, map, 'FAILED LOGIN COUNT', 0);
      }
      usersSheet.getRange(rowNumber, 1, 1, usersSheet.getLastColumn()).setValues([row]);
      const failedUser = userRecordFromRow_(row, map, rowNumber);
      const failedDirectory = directory.users.map(item =>
        item.userId === failedUser.userId ? failedUser : item
      );
      KM_REQUEST_USER_DIRECTORY_ = failedDirectory;
      putCachedJson_('users', failedDirectory, KM_AUTH.USER_DIRECTORY_CACHE_SECONDS);
      writeAuthLogUnsafe_({ identifier: normalizedIdentifier, userId: user.userId, email: user.email,
        action: 'LOGIN', status: 'FAILED', note: 'Invalid credential.', userAgentHash: uaHash });
      if (failedCount >= KM_AUTH.MAX_FAILED_LOGINS) {
        throw new Error('LOGIN_LOCKED: Akun sementara dikunci selama ' + KM_AUTH.LOCK_MINUTES +
          ' menit karena terlalu banyak percobaan gagal.');
      }
      throw new Error('LOGIN_FAILED: email/username atau password tidak valid.');
    }

    setRowValue_(row, map, 'FAILED LOGIN COUNT', 0);
    setRowValue_(row, map, 'LOCKED UNTIL', '');
    setRowValue_(row, map, 'LAST LOGIN AT', now);
    usersSheet.getRange(rowNumber, 1, 1, usersSheet.getLastColumn()).setValues([row]);
    const refreshedUser = userRecordFromRow_(row, map, rowNumber);
    // Perbarui hanya record yang berubah dan pertahankan directory cache.
    const refreshedDirectory = directory.users.map(item =>
      item.userId === refreshedUser.userId ? refreshedUser : item
    );
    KM_REQUEST_USER_DIRECTORY_ = refreshedDirectory;
    putCachedJson_('users', refreshedDirectory, KM_AUTH.USER_DIRECTORY_CACHE_SECONDS);
    const session = createSessionForUser_(refreshedUser, clientInfo);
    const actor = buildActorFromUser_(refreshedUser, 'WEB_APP', session);
    KM_REQUEST_ACTOR_ = actor;
    writeAuthLogUnsafe_({ identifier: normalizedIdentifier, userId: refreshedUser.userId, email: refreshedUser.email,
      action: 'LOGIN', status: 'SUCCESS', note: '', userAgentHash: uaHash, sessionId: session.sessionId });
    authenticated = { session: session, actor: actor, user: refreshedUser };
  } finally {
    lock.releaseLock();
  }

  // Bootstrap hanya membaca konfigurasi dan dijalankan setelah lock dilepas.
  // Mengembalikannya dalam respons login menghemat satu round-trip browser -> Apps Script.
  const mustChangePassword = Boolean(authenticated.user.mustChangePassword);
  return {
    success: true,
    sessionToken: authenticated.session.token,
    user: publicActorContext_(authenticated.actor),
    sessionExpiresAt: authenticated.session.expiresAt,
    mustChangePassword: mustChangePassword,
    bootstrap: mustChangePassword ? null : buildAppBootstrapData_(authenticated.actor)
  };
}

function getSessionState(sessionToken) {
  const actor = getCurrentUserContext_({ token: sessionToken, forceRevalidate: true });
  return {
    authenticated: true,
    user: publicActorContext_(actor),
    bootstrap: actor.mustChangePassword ? null : buildAppBootstrapData_(actor)
  };
}


function getSessionHealth(sessionToken) {
  const actor = validateSessionToken_(sessionToken, { healthCheck: true });
  return {
    authenticated: true,
    serverNow: new Date().toISOString(),
    sessionExpiresAt: actor.sessionExpiresAt || '',
    idleTimeoutMinutes: KM_AUTH.IDLE_TIMEOUT_MINUTES
  };
}

function logoutUser(sessionToken) {
  const parsed = parseSessionToken_(sessionToken);
  if (!parsed) return { success: true };
  const session = readSessionById_(parsed.sessionId);
  if (session && constantTimeEquals_(sha256Hex_(parsed.secret), session.tokenHash)) {
    revokeSessionByRecord_(session, 'LOGOUT');
    const user = readUserDirectory_().find(item => item.userId === session.userId);
    writeAuthLog_({ userId: session.userId, email: user && user.email, action: 'LOGOUT', status: 'SUCCESS', sessionId: session.sessionId });
  }
  KM_REQUEST_ACTOR_ = null;
  return { success: true };
}

function changeOwnPassword(sessionToken, currentPassword, newPassword, clientInfo) {
  const actor = getCurrentUserContext_({ token: sessionToken, forceRevalidate: true, allowPasswordChange: true });
  const users = readUserDirectory_();
  const user = users.find(item => item.userId === actor.userId);
  const currentValue = String(currentPassword || '');
  const nextValue = String(newPassword || '');
  if (!user || currentValue.length > KM_AUTH.PASSWORD_MAX_LENGTH || !verifyPassword_(currentValue, user)) {
    throw new Error('PASSWORD_INVALID: password saat ini tidak sesuai.');
  }
  const errors = passwordPolicyErrors_(nextValue, user);
  if (errors.length) throw new Error('PASSWORD_POLICY: password baru harus memiliki ' + errors.join(', ') + '.');
  // Hindari satu putaran PBKDF2 tambahan: kedua plaintext tersedia hanya pada request ini.
  if (constantTimeEquals_(nextValue, currentValue)) {
    throw new Error('PASSWORD_POLICY: password baru harus berbeda dari password saat ini.');
  }

  // PBKDF2 cukup mahal di Apps Script. Hitung credential sebelum mengambil
  // ScriptLock agar login/update pengguna lain tidak ikut tertahan.
  const credential = createPasswordCredential_(nextValue);
  const lock = acquireScriptLock_(12000, 'perubahan password');
  let response = null;
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.USERS);
    if (!sheet) throw new Error('AUTH_NOT_INITIALIZED: KM_USERS belum tersedia.');
    ensureHeaders_(sheet, KM_APP.USER_HEADERS);
    const map = getHeaderMap_(sheet);
    const row = sheet.getRange(user.rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    const freshUser = userRecordFromRow_(row, map, user.rowNumber);
    if (freshUser.userId !== user.userId || freshUser.sessionVersion !== user.sessionVersion) {
      throw new Error('SESSION_REVOKED: data akun berubah. Silakan login kembali.');
    }

    const now = new Date();
    const nextVersion = freshUser.sessionVersion + 1;
    setRowValue_(row, map, 'PASSWORD HASH', credential.hash);
    setRowValue_(row, map, 'PASSWORD SALT', credential.salt);
    setRowValue_(row, map, 'PASSWORD ITERATIONS', credential.iterations);
    setRowValue_(row, map, 'PASSWORD ALGORITHM', credential.algorithm);
    setRowValue_(row, map, 'PASSWORD CHANGED AT', now);
    setRowValue_(row, map, 'MUST CHANGE PASSWORD', false);
    setRowValue_(row, map, 'FAILED LOGIN COUNT', 0);
    setRowValue_(row, map, 'LOCKED UNTIL', '');
    setRowValue_(row, map, 'SESSION VERSION', nextVersion);
    setRowValue_(row, map, 'UPDATED AT', now);
    setRowValue_(row, map, 'UPDATED BY', actor.email);
    sheet.getRange(user.rowNumber, 1, 1, sheet.getLastColumn()).setValues([row]);
    revokeAllUserSessions_(user.userId, 'PASSWORD_CHANGED');
    invalidateUserDirectoryCache_();

    const refreshed = userRecordFromRow_(row, map, user.rowNumber);
    const session = createSessionForUser_(refreshed, clientInfo);
    const nextActor = buildActorFromUser_(refreshed, 'WEB_APP', session);
    writeAuthLogUnsafe_({ userId: user.userId, email: user.email, action: 'PASSWORD_CHANGED', status: 'SUCCESS', sessionId: session.sessionId });
    KM_REQUEST_ACTOR_ = nextActor;
    response = {
      success: true,
      sessionToken: session.token,
      user: publicActorContext_(nextActor),
      sessionExpiresAt: session.expiresAt,
      bootstrapActor: nextActor
    };
  } finally {
    lock.releaseLock();
  }
  // Jangan menahan lock autentikasi saat membaca konfigurasi bootstrap.
  response.bootstrap = buildAppBootstrapData_(response.bootstrapActor);
  delete response.bootstrapActor;
  return response;
}

function getUserManagementData(sessionToken, forceRefresh) {
  const actor = requirePermission_('MANAGE_USERS', '', { token: sessionToken });
  const users = readUserDirectory_(Boolean(forceRefresh));
  const roleDirectory = readRoleDirectory_(Boolean(forceRefresh));
  const operatingPoints = getManagedOperatingPoints_();
  const publicUsers = users.map(publicUserRecord_);
  return {
    actor: publicActorContext_(actor),
    roles: roleDirectory.filter(role => role.active).map(role => role.role),
    roleDefinitions: roleDirectory,
    permissionDefinitions: Object.keys(KM_RBAC.PERMISSIONS).map(permission => ({
      permission: permission,
      label: KM_PERMISSION_LABELS[permission] || permission
    })),
    operatingPoints: operatingPoints,
    users: publicUsers,
    stats: {
      total: publicUsers.length,
      active: publicUsers.filter(user => user.active).length,
      superAdmin: publicUsers.filter(user => user.active && user.roles.includes('SUPER_ADMIN')).length,
      locked: publicUsers.filter(user => Boolean(user.lockedUntil)).length,
      mustChangePassword: publicUsers.filter(user => Boolean(user.mustChangePassword)).length
    }
  };
}

function saveManagedRole(payload, sessionToken) {
  const actor = requirePermission_('MANAGE_USERS', '', { token: sessionToken });
  if (!payload || typeof payload !== 'object') throw new Error('Data role tidak valid.');
  const role = canonicalRoleCode_(payload.role);
  // Role sistem OH terdiri dari dua karakter. Validasi ini harus sama dengan
  // pattern pada form agar semua role bawaan dapat diedit tanpa pengecualian.
  if (!/^[A-Z][A-Z0-9_]{1,39}$/.test(role)) {
    throw new Error('Kode role harus 2-40 karakter, diawali huruf, dan hanya berisi huruf kapital, angka, atau underscore.');
  }
  const allowedPermissions = new Set(Object.keys(KM_RBAC.PERMISSIONS));
  const permissions = splitAccessValues_(payload.permissions).map(normalizeLookupText_)
    .filter(permission => allowedPermissions.has(permission));
  if (!permissions.length) throw new Error('Pilih minimal satu akses untuk role.');
  if (!permissions.includes('ACCESS_APP')) permissions.unshift('ACCESS_APP');
  const globalScope = false; // Cakupan unit hanya berasal dari KM_USERS.
  const active = payload.active === undefined ? true : boolValue_(payload.active);
  const displayName = safeSheetText_(String(payload.displayName || role.replace(/_/g, ' ')).slice(0, 80));

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = setupRoleDatabase_(ss);
    const map = getHeaderMap_(sheet);
    const values = sheet.getLastRow() < 2 ? [] :
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const matches = values.map((row, index) =>
      canonicalRoleCode_(getRowValue_(row, map, 'ROLE')) === role ? index : -1)
      .filter(index => index >= 0);
    if (matches.length > 1) throw new Error('Role duplikat pada KM_ROLES.');
    const index = matches.length ? matches[0] : -1;
    const existing = index >= 0 ? values[index] : null;
    if (role === 'SUPER_ADMIN' && (!active || !permissions.includes('MANAGE_USERS') || !permissions.includes('ADMIN_SETUP'))) {
      throw new Error('SUPER_ADMIN wajib aktif dan tetap memiliki MANAGE_USERS serta ADMIN_SETUP.');
    }
    if (!active) {
      const assigned = readUserDirectory_(true).some(user => user.active && user.roles.includes(role));
      if (assigned) throw new Error('Role masih digunakan pengguna aktif dan tidak dapat dinonaktifkan.');
    }
    const now = new Date();
    const row = existing ? existing.slice() : new Array(sheet.getLastColumn()).fill('');
    const before = existing ? {
      permissions: splitAccessValues_(getRowValue_(existing, map, 'PERMISSIONS')),
      globalScope: boolValue_(getRowValue_(existing, map, 'GLOBAL SCOPE')),
      active: boolValue_(getRowValue_(existing, map, 'ACTIVE'))
    } : null;
    setRowValue_(row, map, 'ROLE', role);
    setRowValue_(row, map, 'DISPLAY NAME', displayName);
    setRowValue_(row, map, 'PERMISSIONS', permissions.join(', '));
    setRowValue_(row, map, 'GLOBAL SCOPE', globalScope);
    setRowValue_(row, map, 'ACTIVE', active);
    setRowValue_(row, map, 'SYSTEM ROLE', existing ? boolValue_(getRowValue_(existing, map, 'SYSTEM ROLE')) : false);
    setRowValue_(row, map, 'CREATED AT', existing ? getRowValue_(existing, map, 'CREATED AT') : now);
    setRowValue_(row, map, 'CREATED BY', existing ? getRowValue_(existing, map, 'CREATED BY') : actor.email);
    setRowValue_(row, map, 'UPDATED AT', now);
    setRowValue_(row, map, 'UPDATED BY', actor.email);
    const rowNumber = index >= 0 ? index + 2 : sheet.getLastRow() + 1;
    ensureSheetCapacity_(sheet, index >= 0 ? 0 : 1);
    sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
    SpreadsheetApp.flush();
    const storedRow = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    const storedPermissions = splitAccessValues_(getRowValue_(storedRow, map, 'PERMISSIONS'))
      .map(normalizeLookupText_).filter(permission => allowedPermissions.has(permission));
    if (storedPermissions.join('|') !== permissions.join('|')) {
      throw new Error('Penyimpanan permission tidak dapat diverifikasi. Refresh data role lalu coba kembali.');
    }
    invalidateSecurityCaches_();
    logSystemActivity_('', '', index >= 0 ? 'ROLE_UPDATED' : 'ROLE_CREATED', 'KM_ROLES',
      before ? JSON.stringify(before) : '', JSON.stringify({ role, permissions, globalScope, active }),
      displayName, actor, true);
    return { success: true, created: index < 0, role: {
      rowNumber, role, displayName, permissions, globalScope, active,
      systemRole: boolValue_(getRowValue_(storedRow, map, 'SYSTEM ROLE'))
    }};
  } finally {
    lock.releaseLock();
  }
}

function saveManagedUser(payload, sessionToken) {
  const actor = requirePermission_('MANAGE_USERS', '', { token: sessionToken });
  if (!payload || typeof payload !== 'object') throw new Error('Data pengguna tidak valid.');
  const email = normalizeEmail_(payload.email);
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Email pengguna tidak valid.');
  const requestedUsername = normalizeUsername_(payload.username || email.split('@')[0]);
  if (!/^[a-z0-9._-]{3,50}$/.test(requestedUsername)) {
    throw new Error('Username 3-50 karakter dan hanya boleh berisi huruf kecil, angka, titik, underscore, atau dash.');
  }
  const roles = normalizeRoles_(payload.roles);
  if (!roles.length) throw new Error('Pilih minimal satu role.');
  let scopes = splitAccessValues_(payload.operatingPoints);
  if (!scopes.length) throw new Error('Pilih minimal satu Operating Point.');
  if (scopes.includes('*')) {
    if (scopes.length !== 1) throw new Error('Scope semua OP tidak boleh digabung dengan scope tertentu.');
  } else {
    const configuredOperatingPoints = getManagedOperatingPoints_();
    scopes = scopes.map(scope => resolveMasterValue_(configuredOperatingPoints, scope, 'Operating Point scope'));
  }
  const active = payload.active === true || isActiveAccessValue_(payload.active);
  const displayName = safeSheetText_(String(payload.displayName || '').slice(0, 120));

  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(KM_APP.SHEETS.USERS);
    if (!sheet) throw new Error('KM_USERS belum tersedia. Jalankan setup aplikasi terlebih dahulu.');
    ensureHeaders_(sheet, KM_APP.USER_HEADERS);
    const map = getHeaderMap_(sheet);
    const values = sheet.getLastRow() < 2 ? [] : sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

    const matchingIndexes = values.map((row, rowIndex) =>
      normalizeEmail_(getRowValue_(row, map, 'EMAIL')) === email ? rowIndex : -1).filter(rowIndex => rowIndex >= 0);
    if (matchingIndexes.length > 1) throw new Error('Email duplikat pada KM_USERS. Perbaiki sheet sebelum melanjutkan.');
    const index = matchingIndexes.length ? matchingIndexes[0] : -1;
    const usernameConflict = values.some((row, rowIndex) =>
      rowIndex !== index && normalizeUsername_(getRowValue_(row, map, 'USERNAME')) === requestedUsername);
    if (usernameConflict) throw new Error('Username sudah digunakan oleh akun lain.');

    const existing = index >= 0 ? values[index] : null;
    const now = new Date();
    const row = existing ? existing.slice() : new Array(sheet.getLastColumn()).fill('');
    const originalCreatedAt = existing ? getRowValue_(existing, map, 'CREATED AT') : now;
    const originalCreatedBy = existing ? getRowValue_(existing, map, 'CREATED BY') : actor.email;
    const originalRoles = existing ? normalizeRoles_(getRowValue_(existing, map, 'ROLES')) : [];
    const originalActive = existing ? isActiveAccessValue_(getRowValue_(existing, map, 'ACTIVE')) : false;
    const originalScopes = existing ? splitAccessValues_(getRowValue_(existing, map, 'OPERATING POINTS')) : [];
    const userId = existing ? cleanText_(getRowValue_(existing, map, 'USER ID')) || newUserId_() : newUserId_();
    let sessionVersion = existing ? Math.max(1, Number(getRowValue_(existing, map, 'SESSION VERSION')) || 1) : 1;
    let temporaryPassword = '';

    setRowValue_(row, map, 'EMAIL', email);
    setRowValue_(row, map, 'USERNAME', requestedUsername);
    setRowValue_(row, map, 'USER ID', userId);
    setRowValue_(row, map, 'ROLES', roles.join(', '));
    setRowValue_(row, map, 'OPERATING POINTS', scopes.join(', '));
    setRowValue_(row, map, 'ACTIVE', active);
    setRowValue_(row, map, 'DISPLAY NAME', displayName);
    setRowValue_(row, map, 'CREATED AT', originalCreatedAt);
    setRowValue_(row, map, 'CREATED BY', originalCreatedBy);
    setRowValue_(row, map, 'UPDATED AT', now);
    setRowValue_(row, map, 'UPDATED BY', actor.email);

    const passwordMissing = !cleanText_(getRowValue_(row, map, 'PASSWORD HASH')) || !cleanText_(getRowValue_(row, map, 'PASSWORD SALT'));
    if (!existing || passwordMissing) {
      temporaryPassword = generateTemporaryPassword_();
      // Password sementara high-entropy menggunakan verifier HMAC cepat;
      // user wajib menggantinya ke PBKDF2 pada login pertama.
      const credential = createTemporaryPasswordCredential_(temporaryPassword);
      setRowValue_(row, map, 'PASSWORD HASH', credential.hash);
      setRowValue_(row, map, 'PASSWORD SALT', credential.salt);
      setRowValue_(row, map, 'PASSWORD ITERATIONS', credential.iterations);
      setRowValue_(row, map, 'PASSWORD ALGORITHM', credential.algorithm);
      setRowValue_(row, map, 'PASSWORD CHANGED AT', now);
      setRowValue_(row, map, 'MUST CHANGE PASSWORD', true);
      setRowValue_(row, map, 'FAILED LOGIN COUNT', 0);
      setRowValue_(row, map, 'LOCKED UNTIL', '');
    }

    const removingSuperAdmin = existing && originalRoles.includes('SUPER_ADMIN') && (!roles.includes('SUPER_ADMIN') || !active);
    if (removingSuperAdmin) {
      const otherActiveSuperAdmins = values.filter((candidate, rowIndex) => rowIndex !== index &&
        isActiveAccessValue_(getRowValue_(candidate, map, 'ACTIVE')) &&
        normalizeRoles_(getRowValue_(candidate, map, 'ROLES')).includes('SUPER_ADMIN'));
      if (!otherActiveSuperAdmins.length) throw new Error('SUPER_ADMIN aktif terakhir tidak boleh dihapus role-nya atau dinonaktifkan.');
    }

    const securityChanged = Boolean(existing) && (
      originalActive !== active ||
      originalRoles.join('|') !== roles.join('|') ||
      originalScopes.join('|') !== scopes.join('|')
    );
    if (securityChanged) {
      sessionVersion++;
      setRowValue_(row, map, 'SESSION VERSION', sessionVersion);
    } else if (!getRowValue_(row, map, 'SESSION VERSION')) {
      setRowValue_(row, map, 'SESSION VERSION', sessionVersion);
    }

    const targetRow = index >= 0 ? index + 2 : sheet.getLastRow() + 1;
    ensureSheetCapacity_(sheet, index >= 0 ? 0 : 1);
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    if (securityChanged) revokeAllUserSessions_(userId, 'ROLE_SCOPE_OR_STATUS_CHANGED');
    invalidateUserDirectoryCache_();

    // Perubahan user langsung berlaku pada RBAC backend; tidak ada sinkronisasi
    // proteksi Google Sheet yang berjalan di latar belakang.
    logSystemActivity_('', '', index >= 0 ? 'USER_UPDATED' : 'USER_CREATED', 'KM_USERS',
      existing ? JSON.stringify({ email, roles: originalRoles, operatingPoints: originalScopes }) : '',
      JSON.stringify({ email, username: requestedUsername, roles, operatingPoints: scopes, active }),
      'User management', actor, true);
    writeAuthLogUnsafe_({ identifier: requestedUsername, userId, email, action: index >= 0 ? 'USER_UPDATED' : 'USER_CREATED', status: 'SUCCESS', note: actor.email });

    const savedUser = userRecordFromRow_(row, map, targetRow);
    return {
      success: true,
      email,
      username: requestedUsername,
      created: index < 0,
      temporaryPassword,
      mustChangePassword: Boolean(temporaryPassword),
      user: publicUserRecord_(savedUser)
    };
  } finally {
    lock.releaseLock();
  }
}

function setManagedUserActive(email, active, sessionToken) {
  const actor = requirePermission_('MANAGE_USERS', '', { token: sessionToken });
  const targetEmail = normalizeEmail_(email);
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.USERS);
    if (!sheet || sheet.getLastRow() < 2) throw new Error('KM_USERS belum tersedia.');
    ensureHeaders_(sheet, KM_APP.USER_HEADERS);
    const map = getHeaderMap_(sheet);
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const index = values.findIndex(row => normalizeEmail_(getRowValue_(row, map, 'EMAIL')) === targetEmail);
    if (index < 0) throw new Error('Pengguna tidak ditemukan.');
    const nextActive = active === true || isActiveAccessValue_(active);
    const targetRoles = normalizeRoles_(getRowValue_(values[index], map, 'ROLES'));
    if (!nextActive && targetRoles.includes('SUPER_ADMIN')) {
      const otherActiveSuperAdmins = values.filter((row, rowIndex) => rowIndex !== index &&
        isActiveAccessValue_(getRowValue_(row, map, 'ACTIVE')) &&
        normalizeRoles_(getRowValue_(row, map, 'ROLES')).includes('SUPER_ADMIN'));
      if (!otherActiveSuperAdmins.length) throw new Error('SUPER_ADMIN aktif terakhir tidak boleh dinonaktifkan.');
    }
    const rowNumber = index + 2;
    const updatedRow = values[index].slice();
    const userId = cleanText_(getRowValue_(updatedRow, map, 'USER ID'));
    const currentActive = isActiveAccessValue_(getRowValue_(updatedRow, map, 'ACTIVE'));
    if (currentActive === nextActive) {
      return { success: true, email: targetEmail, active: nextActive, unchanged: true, user: publicUserRecord_(userRecordFromRow_(updatedRow, map, rowNumber)) };
    }
    const nextVersion = Math.max(1, Number(getRowValue_(updatedRow, map, 'SESSION VERSION')) || 1) + 1;
    setRowValue_(updatedRow, map, 'ACTIVE', nextActive);
    setRowValue_(updatedRow, map, 'SESSION VERSION', nextVersion);
    setRowValue_(updatedRow, map, 'UPDATED AT', new Date());
    setRowValue_(updatedRow, map, 'UPDATED BY', actor.email);
    sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).setValues([updatedRow]);
    revokeAllUserSessions_(userId, nextActive ? 'ACCOUNT_REACTIVATED' : 'ACCOUNT_DEACTIVATED');
    invalidateUserDirectoryCache_();
    logSystemActivity_('', '', nextActive ? 'USER_ACTIVATED' : 'USER_DEACTIVATED', 'KM_USERS',
      currentActive, nextActive, targetEmail, actor, true);
    return { success: true, email: targetEmail, active: nextActive, user: publicUserRecord_(userRecordFromRow_(updatedRow, map, rowNumber)) };
  } finally {
    lock.releaseLock();
  }
}

function resetManagedUserPassword(email, sessionToken) {
  const actor = requirePermission_('MANAGE_USERS', '', { token: sessionToken });
  const targetEmail = normalizeEmail_(email);
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.USERS);
    if (!sheet || sheet.getLastRow() < 2) throw new Error('KM_USERS belum tersedia.');
    ensureHeaders_(sheet, KM_APP.USER_HEADERS);
    const map = getHeaderMap_(sheet);
    const emailColumn = map['EMAIL'];
    const emails = sheet.getRange(2, emailColumn, sheet.getLastRow() - 1, 1).getDisplayValues();
    const index = emails.findIndex(row => normalizeEmail_(row[0]) === targetEmail);
    if (index < 0) throw new Error('Pengguna tidak ditemukan.');
    const rowNumber = index + 2;
    const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    const userId = cleanText_(getRowValue_(row, map, 'USER ID'));
    const temporaryPassword = generateTemporaryPassword_();
    const credential = createTemporaryPasswordCredential_(temporaryPassword);
    const now = new Date();
    const nextVersion = Math.max(1, Number(getRowValue_(row, map, 'SESSION VERSION')) || 1) + 1;
    setRowValue_(row, map, 'PASSWORD HASH', credential.hash);
    setRowValue_(row, map, 'PASSWORD SALT', credential.salt);
    setRowValue_(row, map, 'PASSWORD ITERATIONS', credential.iterations);
    setRowValue_(row, map, 'PASSWORD ALGORITHM', credential.algorithm);
    setRowValue_(row, map, 'PASSWORD CHANGED AT', now);
    setRowValue_(row, map, 'MUST CHANGE PASSWORD', true);
    setRowValue_(row, map, 'FAILED LOGIN COUNT', 0);
    setRowValue_(row, map, 'LOCKED UNTIL', '');
    setRowValue_(row, map, 'SESSION VERSION', nextVersion);
    setRowValue_(row, map, 'UPDATED AT', now);
    setRowValue_(row, map, 'UPDATED BY', actor.email);
    sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).setValues([row]);
    revokeAllUserSessions_(userId, 'PASSWORD_RESET_BY_ADMIN');
    invalidateUserDirectoryCache_();
    writeAuthLogUnsafe_({ identifier: targetEmail, userId, email: targetEmail, action: 'PASSWORD_RESET', status: 'SUCCESS', note: actor.email });
    logSystemActivity_('', '', 'USER_PASSWORD_RESET', 'KM_USERS', '', '', targetEmail, actor, true);
    return {
      success: true,
      email: targetEmail,
      temporaryPassword,
      mustChangePassword: true,
      user: publicUserRecord_(userRecordFromRow_(row, map, rowNumber))
    };
  } finally {
    lock.releaseLock();
  }
}

function unlockManagedUser(email, sessionToken) {
  const actor = requirePermission_('MANAGE_USERS', '', { token: sessionToken });
  const targetEmail = normalizeEmail_(email);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.USERS);
    if (!sheet || sheet.getLastRow() < 2) throw new Error('KM_USERS belum tersedia.');
    ensureHeaders_(sheet, KM_APP.USER_HEADERS);
    const map = getHeaderMap_(sheet);
    const emails = sheet.getRange(2, map['EMAIL'], sheet.getLastRow() - 1, 1).getDisplayValues();
    const index = emails.findIndex(row => normalizeEmail_(row[0]) === targetEmail);
    if (index < 0) throw new Error('Pengguna tidak ditemukan.');
    const rowNumber = index + 2;
    const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    setRowValue_(row, map, 'FAILED LOGIN COUNT', 0);
    setRowValue_(row, map, 'LOCKED UNTIL', '');
    setRowValue_(row, map, 'UPDATED AT', new Date());
    setRowValue_(row, map, 'UPDATED BY', actor.email);
    sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).setValues([row]);
    invalidateUserDirectoryCache_();
    writeAuthLogUnsafe_({ identifier: targetEmail, userId: cleanText_(getRowValue_(row, map, 'USER ID')), email: targetEmail,
      action: 'ACCOUNT_UNLOCKED', status: 'SUCCESS', note: actor.email });
    return { success: true, user: publicUserRecord_(userRecordFromRow_(row, map, rowNumber)) };
  } finally {
    lock.releaseLock();
  }
}

function revokeManagedUserSessions(email, sessionToken) {
  const actor = requirePermission_('MANAGE_USERS', '', { token: sessionToken });
  const targetEmail = normalizeEmail_(email);
  const users = readUserDirectory_(true);
  const user = users.find(item => item.email === targetEmail);
  if (!user) throw new Error('Pengguna tidak ditemukan.');
  if (user.userId === actor.userId) throw new Error('Gunakan tombol Logout untuk mencabut sesi akun Anda sendiri.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const count = revokeAllUserSessions_(user.userId, 'REVOKED_BY_SUPER_ADMIN');
    writeAuthLogUnsafe_({ identifier: targetEmail, userId: user.userId, email: targetEmail,
      action: 'SESSIONS_REVOKED', status: 'SUCCESS', note: actor.email + ' | ' + count + ' session(s)' });
    return { success: true, revoked: count };
  } finally {
    lock.releaseLock();
  }
}

function syncRbacProtectionsNow(sessionToken) {
  const actor = requirePermission_('MANAGE_USERS', '', { token: sessionToken });
  const result = syncRbacSheetProtections_();
  logSystemActivity_('', '', 'AUTOMATIC_PROTECTION_CLEANUP', 'DATABASE', '', '',
    'Proteksi otomatis lama dibersihkan; akses aplikasi tetap menggunakan RBAC backend.', actor);
  return result || { success: true };
}

function getAuditTrail(limit, sessionToken) {
  requirePermission_('VIEW_AUDIT', '', { token: sessionToken });
  const rowLimit = Math.max(1, Math.min(Number(limit) || 100, 500));
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.LOG);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const map = getHeaderMap_(sheet);
  const startRow = Math.max(2, sheet.getLastRow() - rowLimit + 1);
  const values = sheet.getRange(startRow, 1, sheet.getLastRow() - startRow + 1, sheet.getLastColumn()).getValues();
  return values.reverse().map(row => ({
    id: displayValue_(getRowValue_(row, map, 'LOG ID')),
    timestamp: formatDateTime_(getRowValue_(row, map, 'TIMESTAMP')),
    nopol: displayValue_(getRowValue_(row, map, 'NOMOR POLISI')),
    email: displayValue_(getRowValue_(row, map, 'USER EMAIL')),
    role: displayValue_(getRowValue_(row, map, 'USER ROLE')),
    source: displayValue_(getRowValue_(row, map, 'SOURCE')),
    action: displayValue_(getRowValue_(row, map, 'ACTION')),
    fieldName: displayValue_(getRowValue_(row, map, 'FIELD NAME')),
    oldValue: displayValue_(getRowValue_(row, map, 'OLD VALUE')),
    newValue: displayValue_(getRowValue_(row, map, 'NEW VALUE')),
    note: displayValue_(getRowValue_(row, map, 'NOTE'))
  }));
}



/* ========================================================================== */
/* ADMINISTRATION CENTER                                                      */
/* ========================================================================== */


// ============================================================================
// 03. DATABASE, HISTORY & CACHE
// ============================================================================
/** Canonical database schema, history lifecycle, backfill, and integrity audit. */
const KM_DATABASE_SCHEMA_VERSION = 23;
const DEFAULT_STATUS_3L_OPTIONS = ['WAITING', 'PROCESS', 'DONE', 'KENDALA'];
const UNIT_OPERATIONAL_STATUSES = ['Standby', 'On Delivery', 'Ready to Depart'];
const UNIT_MANUAL_OPERATIONAL_STATUSES = ['Standby', 'On Delivery'];
const UNIT_KM_FIELDS = {
  statusUnit: 'STATUS UNIT', kmOdometer: 'KM ODOMETER',
  kmPhotoId: 'FOTO ODOMETER ID', kmPhotoAddedAt: 'FOTO DITAMBAHKAN AT', kmPhotoAddedBy: 'FOTO DIUNGGAH OLEH',
  kmDate: 'TANGGAL UPDATE KM', kmUpdatedAt: 'KM UPDATED AT', kmUpdatedBy: 'KM UPDATED BY',
};
const UNIT_MASTER_DISPLAY_FIELDS = {
  calibrationNumber: 'KALIBRASI KE',
  status3L: 'STATUS 3L',
  status3LUpdatedAt: 'STATUS 3L UPDATED AT',
  status3LUpdatedBy: 'STATUS 3L UPDATED BY',
  status3LNote: 'CATATAN / NOTE'
};


/**
 * Schema v22 — urutan header enterprise + clean analytics ledger.
 * Kolom yang tidak dikenal tidak dihapus; standardizer hanya memindahkan
 * kolom canonical ke urutan ini dan membiarkan kolom custom setelahnya.
 */
function masterUnitEnterpriseHeaders_() {
  return [
    'UNIT ID', 'NOMOR POLISI', 'TIPE', 'OPERATING POINT', 'NAMA TOH/WOH',
    'STATUS UNIT',
    'KM ODOMETER', 'FOTO ODOMETER ID', 'FOTO DITAMBAHKAN AT', 'FOTO DIUNGGAH OLEH',
    'TANGGAL UPDATE KM', 'KM UPDATED AT', 'KM UPDATED BY', 'KALIBRASI KE',
    'STATUS 3L', 'STATUS 3L UPDATED AT', 'STATUS 3L UPDATED BY', 'CATATAN / NOTE',
    'MASTER STATUS', 'MASTER CREATED AT', 'MASTER CREATED BY',
    'MASTER UPDATED AT', 'MASTER UPDATED BY', 'LAST MOVEMENT ID'
  ];
}

function enterpriseSheetSchemaSpecs_() {
  return [
    [KM_APP.SHEETS.CONFIG, ['OPERATING POINT','NAMA OPC/WHC','NAMA TOH/WOH','STATUS PROSES 3L','PARAMETER','VALUE']],
    [KM_APP.SHEETS.MASTER_UNIT, masterUnitEnterpriseHeaders_()],
    [KM_APP.SHEETS.LOG, KM_APP.LOG_HEADERS],
    [KM_APP.SHEETS.USERS, KM_APP.USER_HEADERS],
    [KM_APP.SHEETS.ROLES, KM_ROLE_HEADERS],
    [KM_APP.SHEETS.SESSIONS, KM_AUTH.SESSION_HEADERS],
    [KM_APP.SHEETS.AUTH_LOG, KM_AUTH.AUTH_LOG_HEADERS],
    [UNIT_KM_LOG_SHEET, UNIT_KM_LOG_HEADERS],
    [ANALYTICS_CALIBRATION_SHEET, ANALYTICS_CALIBRATION_HEADERS],
    [UNIT_STATUS_LOG_SHEET, UNIT_STATUS_LOG_HEADERS],
    [UNIT_STATUS_3L_LOG_SHEET, UNIT_STATUS_3L_LOG_HEADERS],
    [MASTER_UNIT_MANAGEMENT.MOVEMENT_SHEET, MASTER_UNIT_MANAGEMENT.MOVEMENT_HEADERS],
    [MASTER_UNIT_MANAGEMENT.REQUEST_SHEET, MASTER_UNIT_MANAGEMENT.REQUEST_HEADERS]
  ];
}

function canonicalHeaderOrderReady_(sheet, expectedHeaders) {
  if (!sheet || !sheet.getLastColumn()) return false;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const normalized = headers.map(normalizeHeaderKey_);
  let previous = -1;
  for (let i = 0; i < expectedHeaders.length; i++) {
    const index = normalized.indexOf(normalizeHeaderKey_(expectedHeaders[i]));
    if (index < 0 || index <= previous) return false;
    previous = index;
  }
  return true;
}

function reorderSheetColumnsToCanonical_(sheet, expectedHeaders) {
  if (!sheet || !expectedHeaders || !expectedHeaders.length) return { reordered: false, moved: 0 };
  let moved = 0;
  for (let target = 0; target < expectedHeaders.length; target++) {
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const normalized = headers.map(normalizeHeaderKey_);
    const key = normalizeHeaderKey_(expectedHeaders[target]);
    const matches = normalized.reduce((list, candidate, index) => {
      if (candidate === key) list.push(index);
      return list;
    }, []);
    if (matches.length > 1) throw new Error('Header ganda ' + expectedHeaders[target] + ' di ' + sheet.getName() + '.');
    if (!matches.length) continue;
    const current = matches[0];
    if (current === target) {
      if (headers[current] !== expectedHeaders[target]) sheet.getRange(1, current + 1).setValue(expectedHeaders[target]);
      continue;
    }
    sheet.moveColumns(sheet.getRange(1, current + 1, 1, 1), target + 1);
    sheet.getRange(1, target + 1).setValue(expectedHeaders[target]);
    moved++;
  }
  return { reordered: moved > 0, moved: moved };
}

function standardizeEnterpriseHeaderOrder_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const report = [];
  enterpriseSheetSchemaSpecs_().forEach(spec => {
    const sheet = spreadsheet.getSheetByName(spec[0]);
    if (!sheet) return;
    const expected = spec[1];
    if (sheet.getLastRow() === 0) {
      ensureHeaders_(sheet, expected);
    } else if ([UNIT_KM_LOG_SHEET, UNIT_STATUS_LOG_SHEET, UNIT_STATUS_3L_LOG_SHEET,
        MASTER_UNIT_MANAGEMENT.MOVEMENT_SHEET, MASTER_UNIT_MANAGEMENT.REQUEST_SHEET].includes(spec[0])) {
      ensureHistoryHeaders_(sheet, expected);
    } else {
      ensureHeaders_(sheet, expected);
    }
    const result = reorderSheetColumnsToCanonical_(sheet, expected);
    report.push({ sheet: spec[0], movedColumns: result.moved, ordered: canonicalHeaderOrderReady_(sheet, expected) });
  });
  return report;
}

function activeDatabaseSheetNames_() {
  return enterpriseSheetSchemaSpecs_().map(spec => spec[0]);
}

function auditDatabaseSheetInventory_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const active = new Set(activeDatabaseSheetNames_().map(normalizeHeaderKey_));
  return spreadsheet.getSheets().map(sheet => {
    const name = sheet.getName();
    const key = normalizeHeaderKey_(name);
    let classification = 'UNMANAGED';
    let recommendation = 'Tinjau manual; sistem tidak akan menghapus sheet yang tidak dikenal.';
    if (active.has(key)) {
      classification = 'ACTIVE'; recommendation = 'Dipakai aplikasi aktif.';
    } else if (key === normalizeHeaderKey_(LEGACY_UNIT_KM_LOG_SHEET)) {
      classification = 'LEGACY_MERGE'; recommendation = 'Merge ke HISTORY_KALIBRASI_KM lalu hapus dari workbook aktif setelah backup.';
    } else if (key === normalizeHeaderKey_('KM_PENGAJUAN')) {
      classification = 'LEGACY_ARCHIVE'; recommendation = 'Flow lama; kosong dapat dihapus, berisi data diarsipkan setelah backup.';
    } else if (/^ARCHIVE[_ -]/i.test(name)) {
      classification = 'ARCHIVE'; recommendation = 'Arsip historis; hidden dan tidak dibaca dashboard.';
    }
    return {
      name: name,
      classification: classification,
      recommendation: recommendation,
      rows: Math.max(0, sheet.getLastRow() - 1),
      columns: sheet.getLastColumn(),
      hidden: sheet.isSheetHidden()
    };
  });
}
const UNIT_KM_LOG_SHEET = 'HISTORY_KALIBRASI_KM';
const LEGACY_UNIT_KM_LOG_SHEET = 'KM_UNIT_HISTORY';
const UNIT_KM_LOG_HEADERS = [
  'UPDATE ID', 'TIMESTAMP', 'USER EMAIL', 'NOMOR POLISI', 'KALIBRASI KE', 'TANGGAL KM',
  'STATUS UNIT', 'KM ODOMETER', 'CATATAN', 'DATA SEBELUM', 'FOTO ODOMETER ID',
  'FOTO DITAMBAHKAN AT', 'FOTO DIUNGGAH OLEH',
  // Context siklus dikelompokkan agar audit history mudah dibaca manusia maupun query.
  'OPERATING POINT', 'NAMA TOH/WOH', 'STATUS 3L AWAL', 'STATUS 3L UPDATED AT', 'STATUS 3L UPDATED BY',
  'COMPLETED AT', 'COMPLETED BY',
  'HISTORY ID', 'REQUEST ID', 'CALIBRATION SEQUENCE', 'RECORD TYPE', 'LIFECYCLE STATUS',
  'SOURCE', 'REASON', 'DATA SETELAH'
];
const UNIT_KM_HISTORY_CONTEXT_HEADERS = Object.freeze([
  'OPERATING POINT', 'NAMA TOH/WOH', 'STATUS 3L AWAL',
  'STATUS 3L UPDATED AT', 'STATUS 3L UPDATED BY', 'COMPLETED AT', 'COMPLETED BY'
]);
const ANALYTICS_CALIBRATION_SHEET = 'ANALYTICS_KALIBRASI_KM';
const ANALYTICS_CALIBRATION_HEADERS = Object.freeze([
  'UPDATE ID', 'TIMESTAMP', 'USER EMAIL', 'OPERATING POINT', 'NAMA TOH/WOH',
  'NOMOR POLISI', 'STATUS UNIT', 'KALIBRASI KE', 'TANGGAL KM', 'KM ODOMETER',
  'FOTO ODOMETER ID', 'FOTO DITAMBAHKAN AT', 'FOTO DIUNGGAH OLEH',
  'STATUS 3L AWAL', 'STATUS 3L', 'STATUS 3L UPDATED BY', 'STATUS 3L UPDATED AT',
  'COMPLETED BY', 'COMPLETED AT', 'HISTORY ID', 'REQUEST ID', 'CALIBRATION SEQUENCE'
]);


/**
 * Clean analytics ledger. One row represents exactly one calibration request.
 * HISTORY_KALIBRASI_KM remains the append-only audit/event trail; this sheet is
 * an application-managed projection used by Analisa & Trend.
 */
function analyticsCalibrationKey_(nopol, calibrationNumber) {
  const plate = normalizeNopol_(nopol);
  const sequence = parseCalibrationNumber_(calibrationNumber);
  return plate && sequence ? plate + '|' + sequence : '';
}

function ensureAnalyticsCalibrationSheet_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(spreadsheet, ANALYTICS_CALIBRATION_SHEET);
  ensureHistoryHeaders_(sheet, ANALYTICS_CALIBRATION_HEADERS);
  sheet.setFrozenRows(1);
  return sheet;
}

function analyticsCalibrationOpeningScore_(row, map) {
  const reason = normalizeLookupText_(getRowValue_(row, map, 'REASON'));
  const source = normalizeLookupText_(getRowValue_(row, map, 'SOURCE'));
  const user = normalizeLookupText_(getRowValue_(row, map, 'USER EMAIL'));
  let score = 0;
  if (reason !== 'LEGACY_MASTER_SNAPSHOT') score += 1000;
  if (reason === 'CALIBRATION_CREATED') score += 300;
  if (source === 'DASHBOARD' || source === 'SPREADSHEET') score += 100;
  if (user && user !== 'SYSTEM') score += 50;
  if (cleanText_(getRowValue_(row, map, 'REQUEST ID'))) score += 25;
  if (normalizeOdometerPhotoId_(historyOdometerPhotoId_(row, map))) score += 25;
  return score;
}

function analyticsDateMillis_(value) {
  const date = parseDateTimeInput_(value) || parseDateInput_(value);
  return date instanceof Date && !isNaN(date.getTime()) ? date.getTime() : 0;
}

function buildAnalyticsCalibrationProjection_(ss, requestedKeys) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const keyFilter = requestedKeys && requestedKeys.size ? requestedKeys : null;
  const history = spreadsheet.getSheetByName(UNIT_KM_LOG_SHEET);
  if (!history || history.getLastRow() < 2) return [];
  const historyMap = getHistoryHeaderMap_(history, UNIT_KM_LOG_HEADERS);
  const historyRows = history.getRange(2, 1, history.getLastRow() - 1, history.getLastColumn()).getValues();
  const openings = Object.create(null);
  const finals = Object.create(null);

  historyRows.forEach((row, offset) => {
    const nopol = normalizeDisplayNopol_(getRowValue_(row, historyMap, 'NOMOR POLISI'));
    const sequence = parseCalibrationNumber_(getRowValue_(row, historyMap, 'KALIBRASI KE') ||
      getRowValue_(row, historyMap, 'CALIBRATION SEQUENCE'));
    const key = analyticsCalibrationKey_(nopol, sequence);
    if (!key || (keyFilter && !keyFilter.has(key))) return;
    const type = normalizeLookupText_(getRowValue_(row, historyMap, 'RECORD TYPE'));
    const lifecycle = normalizeLookupText_(getRowValue_(row, historyMap, 'LIFECYCLE STATUS'));
    if (type === 'CALIBRATION_OPENED') {
      const candidate = {
        row: row,
        rowNumber: offset + 2,
        score: analyticsCalibrationOpeningScore_(row, historyMap),
        timestamp: analyticsDateMillis_(getRowValue_(row, historyMap, 'TIMESTAMP'))
      };
      const current = openings[key];
      if (!current || candidate.score > current.score ||
          (candidate.score === current.score && candidate.timestamp && (!current.timestamp || candidate.timestamp < current.timestamp)) ||
          (candidate.score === current.score && candidate.timestamp === current.timestamp && candidate.rowNumber > current.rowNumber)) {
        openings[key] = candidate;
      }
      return;
    }
    if (type === 'FINAL_SNAPSHOT' && lifecycle === 'LOCKED') {
      const completedAt = getRowValue_(row, historyMap, 'COMPLETED AT') || getRowValue_(row, historyMap, 'TIMESTAMP');
      const candidate = { row: row, rowNumber: offset + 2, timestamp: analyticsDateMillis_(completedAt) };
      const current = finals[key];
      if (!current || (candidate.timestamp && (!current.timestamp || candidate.timestamp < current.timestamp)) ||
          (candidate.timestamp === current.timestamp && candidate.rowNumber > current.rowNumber)) {
        finals[key] = candidate;
      }
    }
  });

  const latest3L = Object.create(null);
  const status3LLog = spreadsheet.getSheetByName(UNIT_STATUS_3L_LOG_SHEET);
  if (status3LLog && status3LLog.getLastRow() > 1) {
    const map = getHistoryHeaderMap_(status3LLog, UNIT_STATUS_3L_LOG_HEADERS);
    const rows = status3LLog.getRange(2, 1, status3LLog.getLastRow() - 1, status3LLog.getLastColumn()).getValues();
    rows.forEach((row, offset) => {
      const nopol = normalizeDisplayNopol_(getRowValue_(row, map, 'NOMOR POLISI'));
      const sequence = parseCalibrationNumber_(getRowValue_(row, map, 'KALIBRASI KE'));
      const key = analyticsCalibrationKey_(nopol, sequence);
      if (!key || !openings[key] || (keyFilter && !keyFilter.has(key))) return;
      const timestamp = analyticsDateMillis_(getRowValue_(row, map, 'TIMESTAMP'));
      const status3L = displayValue_(getRowValue_(row, map, 'STATUS BARU'));
      const source = displayValue_(getRowValue_(row, map, 'SOURCE'));
      const reason = map['REASON'] ? displayValue_(getRowValue_(row, map, 'REASON')) : '';
      const rawActor = getRowValue_(row, map, 'USER EMAIL');
      // Non-DONE legacy state can still be useful as a current projection, but DONE
      // is a completion claim and therefore must have independent trusted evidence.
      if (normalizeLookupText_(status3L) === 'DONE' &&
          !isTrustedStatus3LEvent_(rawActor, source, reason, status3L)) return;
      const current = latest3L[key];
      if (!current || timestamp > current.timestamp || (timestamp === current.timestamp && offset > current.offset)) {
        latest3L[key] = {
          timestamp: timestamp,
          offset: offset,
          status3L: status3L,
          updatedAt: getRowValue_(row, map, 'TIMESTAMP'),
          updatedBy: normalizeDoneStatus3LActor_(status3L, rawActor),
          statusUnit: displayValue_(getRowValue_(row, map, 'STATUS UNIT SESUDAH'))
        };
      }
    });
  }

  const latestUnit = Object.create(null);
  const statusUnitLog = spreadsheet.getSheetByName(UNIT_STATUS_LOG_SHEET);
  if (statusUnitLog && statusUnitLog.getLastRow() > 1) {
    const map = getHistoryHeaderMap_(statusUnitLog, UNIT_STATUS_LOG_HEADERS);
    const rows = statusUnitLog.getRange(2, 1, statusUnitLog.getLastRow() - 1, statusUnitLog.getLastColumn()).getValues();
    rows.forEach((row, offset) => {
      const nopol = normalizeDisplayNopol_(getRowValue_(row, map, 'NOMOR POLISI'));
      const sequence = parseCalibrationNumber_(getRowValue_(row, map, 'KALIBRASI KE'));
      const key = analyticsCalibrationKey_(nopol, sequence);
      if (!key || !openings[key] || (keyFilter && !keyFilter.has(key))) return;
      const timestamp = analyticsDateMillis_(getRowValue_(row, map, 'TIMESTAMP'));
      const current = latestUnit[key];
      if (!current || timestamp > current.timestamp || (timestamp === current.timestamp && offset > current.offset)) {
        latestUnit[key] = {
          timestamp: timestamp,
          offset: offset,
          statusUnit: displayValue_(getRowValue_(row, map, 'STATUS BARU'))
        };
      }
    });
  }

  const masterByPlate = Object.create(null);
  const master = spreadsheet.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  if (master && master.getLastRow() > 1) {
    const headers = master.getRange(1, 1, 1, master.getLastColumn()).getDisplayValues()[0];
    const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
    const rows = master.getRange(2, 1, master.getLastRow() - 1, master.getLastColumn()).getValues();
    rows.forEach(row => {
      const nopol = indexes.nopol >= 0 ? normalizeDisplayNopol_(row[indexes.nopol]) : '';
      if (nopol) masterByPlate[normalizeNopol_(nopol)] = { row: row, indexes: indexes };
    });
  }

  const output = [];
  Object.keys(openings).forEach(key => {
    const open = openings[key].row;
    const nopol = normalizeDisplayNopol_(getRowValue_(open, historyMap, 'NOMOR POLISI'));
    const sequence = parseCalibrationNumber_(getRowValue_(open, historyMap, 'KALIBRASI KE') ||
      getRowValue_(open, historyMap, 'CALIBRATION SEQUENCE'));
    const final = finals[key] ? finals[key].row : null;
    const latest = latest3L[key] || null;
    const unitLatest = latestUnit[key] || null;
    const masterEntry = masterByPlate[normalizeNopol_(nopol)] || null;
    const masterIsSameCycle = Boolean(masterEntry && masterEntry.indexes.calibrationNumber >= 0 &&
      parseCalibrationNumber_(masterEntry.row[masterEntry.indexes.calibrationNumber]) === sequence);

    // Every calibration request starts in WAITING. Do not inherit backfilled legacy
    // STATUS 3L values from raw history because older repair/migration rows may contain
    // the *current* state instead of the state at request creation.
    let status3L = latest ? latest.status3L : safeWaitingStatus3LForHistory_();
    let status3LUpdatedAt = latest ? latest.updatedAt :
      (getRowValue_(open, historyMap, 'STATUS 3L UPDATED AT') || getRowValue_(open, historyMap, 'TIMESTAMP'));
    let status3LUpdatedBy = latest ? latest.updatedBy : 'SYSTEM';
    let statusUnit = unitLatest && unitLatest.statusUnit
      ? unitLatest.statusUnit : displayValue_(getRowValue_(open, historyMap, 'STATUS UNIT'));

    // A historical request that has already been superseded by a newer calibration
    // must no longer follow later operational movement of the unit. Freeze it at its
    // request-opening state unless a valid FINAL_SNAPSHOT exists. Legacy openings that
    // incorrectly carried Ready to Depart are normalized to Standby because a new
    // calibration request is created only while the unit is eligible for calibration.
    if (!final && !masterIsSameCycle) {
      const openingStatusUnit = displayValue_(getRowValue_(open, historyMap, 'STATUS UNIT'));
      statusUnit = normalizeLookupText_(openingStatusUnit) === 'READY TO DEPART'
        ? 'Standby' : (openingStatusUnit || 'Standby');
      status3L = safeWaitingStatus3LForHistory_();
      status3LUpdatedAt = getRowValue_(open, historyMap, 'TIMESTAMP') ||
        getRowValue_(open, historyMap, 'STATUS 3L UPDATED AT');
      status3LUpdatedBy = 'SYSTEM';
    }

    if (masterIsSameCycle) {
      const m = masterEntry;
      const masterStatusAt = m.indexes.status3LUpdatedAt >= 0 ? m.row[m.indexes.status3LUpdatedAt] : '';
      const masterStatusMillis = analyticsDateMillis_(masterStatusAt);
      const latestMillis = latest ? latest.timestamp : 0;
      // Live primary state may precede its asynchronous edit ledger. Display it
      // for the current cycle; historical completion still uses trusted snapshots.
      if (cleanText_(m.row[m.indexes.status3L]) || !latest || masterStatusMillis >= latestMillis) {
        const masterStatus3L = m.indexes.status3L >= 0 ? displayValue_(m.row[m.indexes.status3L]) : status3L;
        status3L = masterStatus3L;
        status3LUpdatedAt = masterStatusAt || status3LUpdatedAt;
        status3LUpdatedBy = m.indexes.status3LUpdatedBy >= 0
          ? normalizeDoneStatus3LActor_(status3L, m.row[m.indexes.status3LUpdatedBy]) : status3LUpdatedBy;
      }
      statusUnit = m.indexes.statusUnit >= 0 ? displayValue_(m.row[m.indexes.statusUnit]) : statusUnit;
    }

    let completedAt = '';
    let completedBy = '';
    if (final) {
      status3L = 'DONE';
      statusUnit = 'Ready to Depart';
      completedAt = getRowValue_(final, historyMap, 'COMPLETED AT') || getRowValue_(final, historyMap, 'TIMESTAMP');
      const finalActor = preferredDoneStatus3LActor_(
        getRowValue_(final, historyMap, 'STATUS 3L UPDATED BY'),
        getRowValue_(final, historyMap, 'COMPLETED BY') || getRowValue_(final, historyMap, 'USER EMAIL')
      );
      completedBy = finalActor;
      status3LUpdatedAt = getRowValue_(final, historyMap, 'STATUS 3L UPDATED AT') || completedAt;
      status3LUpdatedBy = finalActor;
    } else if (normalizeLookupText_(status3L) === 'DONE') {
      statusUnit = 'Ready to Depart';
      status3LUpdatedBy = preferredDoneStatus3LActor_(status3LUpdatedBy, '');
      completedAt = status3LUpdatedAt;
      completedBy = status3LUpdatedBy;
    }

    const historyId = cleanText_(getRowValue_(open, historyMap, 'HISTORY ID'));
    const updateId = cleanText_(getRowValue_(open, historyMap, 'UPDATE ID'));
    const requestId = cleanText_(getRowValue_(open, historyMap, 'REQUEST ID')) || updateId;
    const initialStatus = safeWaitingStatus3LForHistory_();
    const km = parseUnitKm_(getRowValue_(open, historyMap, 'KM ODOMETER'));
    output.push([
      updateId,
      getRowValue_(open, historyMap, 'TIMESTAMP'),
      safeSheetText_(getRowValue_(open, historyMap, 'USER EMAIL')),
      safeSheetText_(getRowValue_(open, historyMap, 'OPERATING POINT')),
      safeSheetText_(getRowValue_(open, historyMap, 'NAMA TOH/WOH')),
      safeSheetText_(nopol),
      safeSheetText_(statusUnit),
      sequence,
      getRowValue_(open, historyMap, 'TANGGAL KM'),
      km == null ? '' : km,
      normalizeOdometerPhotoId_(historyOdometerPhotoId_(open, historyMap)),
      getRowValue_(open, historyMap, 'FOTO DITAMBAHKAN AT'),
      safeSheetText_(getRowValue_(open, historyMap, 'FOTO DIUNGGAH OLEH')),
      safeSheetText_(initialStatus),
      safeSheetText_(status3L),
      safeSheetText_(normalizeDoneStatus3LActor_(status3L, status3LUpdatedBy)),
      status3LUpdatedAt,
      safeSheetText_(completedBy),
      completedAt,
      safeSheetText_(historyId),
      safeSheetText_(requestId),
      sequence
    ]);
  });

  output.sort((a, b) => {
    const ta = analyticsDateMillis_(a[1]);
    const tb = analyticsDateMillis_(b[1]);
    if (ta !== tb) return ta - tb;
    const plateCompare = normalizeNopol_(a[5]).localeCompare(normalizeNopol_(b[5]));
    return plateCompare || Number(a[21] || 0) - Number(b[21] || 0);
  });
  return output;
}

function analyticsCalibrationIndex_(sheet) {
  const result = { byKey: Object.create(null), duplicates: [] };
  if (!sheet || sheet.getLastRow() < 2) return result;
  const values = sheet.getRange(2, 6, sheet.getLastRow() - 1, 17).getValues();
  values.forEach((row, offset) => {
    // F:NOMOR POLISI => offset 0, H:KALIBRASI KE => offset 2, V:CALIBRATION SEQUENCE => offset 16.
    const key = analyticsCalibrationKey_(row[0], row[16] || row[2]);
    if (!key) return;
    if (result.byKey[key]) result.duplicates.push({ key: key, rows: [result.byKey[key], offset + 2] });
    else result.byKey[key] = offset + 2;
  });
  return result;
}

function upsertAnalyticsCalibrationRows_(rows) {
  const values = Array.isArray(rows) ? rows.filter(row => Array.isArray(row) && row.length) : [];
  if (!values.length) return { inserted: 0, updated: 0 };
  const sheet = ensureAnalyticsCalibrationSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const index = analyticsCalibrationIndex_(sheet);
  if (index.duplicates.length) throw new Error('ANALYTICS_DUPLICATE_KEY: jalankan Rebuild Analytics Ledger.');
  let inserted = 0;
  let updated = 0;
  values.forEach(row => {
    const key = analyticsCalibrationKey_(row[5], row[21] || row[7]);
    if (!key) return;
    const targetRow = index.byKey[key];
    if (targetRow) {
      sheet.getRange(targetRow, 1, 1, ANALYTICS_CALIBRATION_HEADERS.length)
        .setValues([row.slice(0, ANALYTICS_CALIBRATION_HEADERS.length)]);
      updated++;
    } else {
      ensureSheetCapacity_(sheet, 1);
      const nextRow = sheet.getLastRow() + 1;
      sheet.getRange(nextRow, 1, 1, ANALYTICS_CALIBRATION_HEADERS.length)
        .setValues([row.slice(0, ANALYTICS_CALIBRATION_HEADERS.length)]);
      index.byKey[key] = nextRow;
      inserted++;
    }
  });
  return { inserted: inserted, updated: updated };
}

function updateAnalyticsCalibrationState_(entries) {
  const states = Array.isArray(entries) ? entries.filter(Boolean) : [];
  if (!states.length) return { updated: 0, rebuiltMissing: 0 };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureAnalyticsCalibrationSheet_(ss);
  let index = analyticsCalibrationIndex_(sheet);
  if (index.duplicates.length) throw new Error('ANALYTICS_DUPLICATE_KEY: jalankan Rebuild Analytics Ledger.');

  const missingKeys = new Set();
  states.forEach(entry => {
    const key = analyticsCalibrationKey_(entry.nopol, entry.calibrationNumber);
    if (key && !index.byKey[key]) missingKeys.add(key);
  });
  if (missingKeys.size) {
    upsertAnalyticsCalibrationRows_(buildAnalyticsCalibrationProjection_(ss, missingKeys));
    index = analyticsCalibrationIndex_(sheet);
  }

  // Completed analytics rows are immutable lifecycle snapshots. Operational changes
  // after completion (for example Ready to Depart -> On Delivery, which also clears
  // Status 3L in MASTER_UNIT) must not erase the calibration completion from trend data.
  // A deliberate integrity correction can still rebuild the ledger from canonical history.
  const targetRows = Array.from(new Set(states.map(entry => {
    const key = analyticsCalibrationKey_(entry.nopol, entry.calibrationNumber);
    return key ? Number(index.byKey[key] || 0) : 0;
  }).filter(Boolean)));
  const lockedCompletedRows = new Set();
  if (targetRows.length) {
    const ranges = sheet.getRangeList(targetRows.map(rowNumber => 'O' + rowNumber + ':S' + rowNumber)).getRanges();
    ranges.forEach((range, i) => {
      const values = range.getValues()[0] || [];
      if (normalizeLookupText_(values[0]) === 'DONE' && values[4]) lockedCompletedRows.add(targetRows[i]);
    });
  }

  const groups = {
    statusUnit: Object.create(null), status3L: Object.create(null), actor: Object.create(null),
    updatedAt: Object.create(null), completedBy: Object.create(null), completedAt: Object.create(null)
  };
  const blankCompletedBy = [];
  const blankCompletedAt = [];
  let updated = 0;
  let skippedLocked = 0;
  const groupAddress = (bucket, key, address, value) => {
    if (!bucket[key]) bucket[key] = { value: value, addresses: [] };
    bucket[key].addresses.push(address);
  };

  states.forEach(entry => {
    const key = analyticsCalibrationKey_(entry.nopol, entry.calibrationNumber);
    const rowNumber = key ? index.byKey[key] : 0;
    if (!rowNumber) return;
    const status3L = displayValue_(entry.status3L);
    const done = normalizeLookupText_(status3L) === 'DONE';
    if (lockedCompletedRows.has(rowNumber) && !done) {
      skippedLocked++;
      return;
    }
    const actor = normalizeDoneStatus3LActor_(status3L, entry.status3LUpdatedBy);
    const statusUnit = done ? 'Ready to Depart' : displayValue_(entry.statusUnit);
    const updatedAt = entry.status3LUpdatedAt || '';
    const completedAt = done ? (entry.completedAt || updatedAt || '') : '';

    groupAddress(groups.statusUnit, normalizeLookupText_(statusUnit), 'G' + rowNumber, safeSheetText_(statusUnit));
    groupAddress(groups.status3L, normalizeLookupText_(status3L), 'O' + rowNumber, safeSheetText_(status3L));
    groupAddress(groups.actor, normalizeLookupText_(actor), 'P' + rowNumber, safeSheetText_(actor));
    const updatedKey = updatedAt instanceof Date ? String(updatedAt.getTime()) : cleanText_(updatedAt);
    groupAddress(groups.updatedAt, updatedKey, 'Q' + rowNumber, updatedAt);
    if (done) {
      groupAddress(groups.completedBy, normalizeLookupText_(actor), 'R' + rowNumber, safeSheetText_(actor));
      const completedKey = completedAt instanceof Date ? String(completedAt.getTime()) : cleanText_(completedAt);
      groupAddress(groups.completedAt, completedKey, 'S' + rowNumber, completedAt);
    } else {
      blankCompletedBy.push('R' + rowNumber);
      blankCompletedAt.push('S' + rowNumber);
    }
    updated++;
  });

  Object.keys(groups).forEach(groupName => {
    Object.keys(groups[groupName]).forEach(key => {
      const item = groups[groupName][key];
      if (item.addresses.length) sheet.getRangeList(item.addresses).setValue(item.value);
    });
  });
  if (blankCompletedBy.length) sheet.getRangeList(blankCompletedBy).clearContent();
  if (blankCompletedAt.length) sheet.getRangeList(blankCompletedAt).clearContent();
  return { updated: updated, rebuiltMissing: missingKeys.size, skippedLocked: skippedLocked };
}

function analyticsStateFromMasterRow_(row, indexes) {
  if (!row || !indexes || indexes.nopol < 0 || indexes.calibrationNumber < 0) return null;
  const nopol = normalizeDisplayNopol_(row[indexes.nopol]);
  const calibrationNumber = parseCalibrationNumber_(row[indexes.calibrationNumber]);
  if (!nopol || !calibrationNumber) return null;
  const status3L = indexes.status3L >= 0 ? displayValue_(row[indexes.status3L]) : '';
  const status3LUpdatedAt = indexes.status3LUpdatedAt >= 0 ? row[indexes.status3LUpdatedAt] : '';
  return {
    nopol: nopol,
    calibrationNumber: calibrationNumber,
    statusUnit: indexes.statusUnit >= 0 ? displayValue_(row[indexes.statusUnit]) : '',
    status3L: status3L,
    status3LUpdatedAt: status3LUpdatedAt,
    status3LUpdatedBy: indexes.status3LUpdatedBy >= 0
      ? normalizeDoneStatus3LActor_(status3L, row[indexes.status3LUpdatedBy]) : '',
    completedAt: normalizeLookupText_(status3L) === 'DONE' ? status3LUpdatedAt : ''
  };
}

function rebuildAnalyticsCalibrationLedgerCore_(actor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureAnalyticsCalibrationSheet_(ss);
  const rows = buildAnalyticsCalibrationProjection_(ss);
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  if (rows.length) {
    ensureSheetCapacity_(sheet, rows.length);
    sheet.getRange(2, 1, rows.length, ANALYTICS_CALIBRATION_HEADERS.length).setValues(rows);
  }
  if (sheet.getMaxColumns() > ANALYTICS_CALIBRATION_HEADERS.length) {
    sheet.deleteColumns(ANALYTICS_CALIBRATION_HEADERS.length + 1,
      sheet.getMaxColumns() - ANALYTICS_CALIBRATION_HEADERS.length);
  }
  trimSheetGrid_(sheet, ANALYTICS_CALIBRATION_HEADERS.length, 100, 101);
  formatDatabaseTables_(ss);
  SpreadsheetApp.flush();
  invalidateAppCaches_();
  if (actor) runNonCritical_('audit rebuild analytics ledger', () => logSystemActivity_('', '',
    'ANALYTICS_LEDGER_REBUILT', ANALYTICS_CALIBRATION_SHEET, '', String(rows.length),
    'Clean ledger dibangun ulang dari evidence history yang canonical.', actor, true));
  return { success: true, rows: rows.length, sheet: ANALYTICS_CALIBRATION_SHEET };
}

function rebuildAnalyticsCalibrationLedger_(actor) {
  const lock = acquireScriptLock_(60000, 'rebuild analytics calibration ledger');
  try {
    return rebuildAnalyticsCalibrationLedgerCore_(actor);
  } finally {
    lock.releaseLock();
  }
}

function rebuildAnalyticsCalibrationLedger(sessionToken) {
  const actor = requirePermission_('ADMIN_SETUP', '', { token: sessionToken });
  return rebuildAnalyticsCalibrationLedger_(actor);
}

function rebuildAnalyticsCalibrationLedgerFromSheet() {
  return rebuildAnalyticsCalibrationLedger_(requireSheetMaintenanceAdmin_());
}

function analyticsComparableValue_(header, value) {
  const key = normalizeHeaderKey_(header);
  if (/TIMESTAMP| AT$/.test(key) || key === 'TANGGAL KM') {
    return analyticsDateMillis_(value) || '';
  }
  if (key === 'KM ODOMETER' || key === 'KALIBRASI KE' || key === 'CALIBRATION SEQUENCE') {
    return value === '' || value == null ? '' : Number(value);
  }
  if (key === 'NOMOR POLISI') return normalizeNopol_(value);
  if (key === 'STATUS 3L UPDATED BY' || key === 'COMPLETED BY') return normalizeLookupText_(value);
  return cleanText_(value);
}

function auditAnalyticsCalibrationLedger_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ANALYTICS_CALIBRATION_SHEET);
  const report = {
    ready: false,
    sheetExists: Boolean(sheet),
    rows: 0,
    expectedRows: 0,
    duplicateKeys: [],
    duplicateRequestIds: [],
    missingKeys: [],
    extraKeys: [],
    mismatches: [],
    blankRequestId: 0,
    blankPhoto: 0,
    blankOperatingPoint: 0,
    blankToh: 0,
    initialStatusNotWaiting: 0,
    doneNotReady: 0,
    doneNotTeam3L: 0,
    doneCompletedByNotTeam3L: 0,
    doneMissingCompletedAt: 0,
    nonDoneHasCompletion: 0,
    protectionMode: 'BACKEND_RBAC_ONLY'
  };
  if (!sheet) return report;

  const map = getHistoryHeaderMap_(sheet, ANALYTICS_CALIBRATION_HEADERS);
  const rows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues() : [];
  report.rows = rows.length;
  const actual = Object.create(null);
  const requestIds = Object.create(null);
  rows.forEach((row, offset) => {
    const key = analyticsCalibrationKey_(getRowValue_(row, map, 'NOMOR POLISI'),
      getRowValue_(row, map, 'CALIBRATION SEQUENCE') || getRowValue_(row, map, 'KALIBRASI KE'));
    if (!key) return;
    if (actual[key] && report.duplicateKeys.length < 100) {
      report.duplicateKeys.push({ key: key, rows: [actual[key].rowNumber, offset + 2] });
    }
    actual[key] = { row: row, rowNumber: offset + 2 };
    const requestId = cleanText_(getRowValue_(row, map, 'REQUEST ID'));
    if (!requestId) report.blankRequestId++;
    else if (requestIds[requestId]) {
      if (report.duplicateRequestIds.length < 100) {
        report.duplicateRequestIds.push({ requestId: requestId, rows: [requestIds[requestId], offset + 2] });
      }
    } else requestIds[requestId] = offset + 2;
    if (!normalizeOdometerPhotoId_(getRowValue_(row, map, 'FOTO ODOMETER ID'))) report.blankPhoto++;
    if (!cleanText_(getRowValue_(row, map, 'OPERATING POINT'))) report.blankOperatingPoint++;
    if (!cleanText_(getRowValue_(row, map, 'NAMA TOH/WOH'))) report.blankToh++;
    if (normalizeLookupText_(getRowValue_(row, map, 'STATUS 3L AWAL')) !== normalizeLookupText_(safeWaitingStatus3LForHistory_())) {
      report.initialStatusNotWaiting++;
    }
    if (normalizeLookupText_(getRowValue_(row, map, 'STATUS 3L')) === 'DONE') {
      if (normalizeLookupText_(getRowValue_(row, map, 'STATUS UNIT')) !== normalizeLookupText_('Ready to Depart')) {
        report.doneNotReady++;
      }
      if (normalizeLookupText_(getRowValue_(row, map, 'STATUS 3L UPDATED BY')) !==
          normalizeLookupText_(KM_STATUS_3L_SHEET_FALLBACK)) {
        report.doneNotTeam3L++;
      }
      if (normalizeLookupText_(getRowValue_(row, map, 'COMPLETED BY')) !==
          normalizeLookupText_(KM_STATUS_3L_SHEET_FALLBACK)) {
        report.doneCompletedByNotTeam3L++;
      }
      if (!getRowValue_(row, map, 'COMPLETED AT')) report.doneMissingCompletedAt++;
    } else if (getRowValue_(row, map, 'COMPLETED AT') || cleanText_(getRowValue_(row, map, 'COMPLETED BY'))) {
      report.nonDoneHasCompletion++;
    }
  });

  const expectedRows = buildAnalyticsCalibrationProjection_(ss);
  report.expectedRows = expectedRows.length;
  const expected = Object.create(null);
  expectedRows.forEach(row => {
    expected[analyticsCalibrationKey_(row[5], row[21] || row[7])] = row;
  });

  Object.keys(expected).forEach(key => {
    if (!actual[key]) {
      if (report.missingKeys.length < 100) report.missingKeys.push(key);
      return;
    }
    const actualRow = actual[key].row;
    const expectedRow = expected[key];
    for (let i = 0; i < ANALYTICS_CALIBRATION_HEADERS.length; i++) {
      if (analyticsComparableValue_(ANALYTICS_CALIBRATION_HEADERS[i], actualRow[i]) !==
          analyticsComparableValue_(ANALYTICS_CALIBRATION_HEADERS[i], expectedRow[i])) {
        if (report.mismatches.length < 100) {
          report.mismatches.push({
            key: key,
            row: actual[key].rowNumber,
            field: ANALYTICS_CALIBRATION_HEADERS[i],
            current: displayValue_(actualRow[i]),
            expected: displayValue_(expectedRow[i])
          });
        }
        break;
      }
    }
  });
  Object.keys(actual).forEach(key => {
    if (!expected[key] && report.extraKeys.length < 100) report.extraKeys.push(key);
  });

  report.ready = !report.duplicateKeys.length && !report.duplicateRequestIds.length &&
    !report.missingKeys.length && !report.extraKeys.length && !report.mismatches.length &&
    !report.blankRequestId && !report.blankPhoto && !report.blankOperatingPoint && !report.blankToh &&
    !report.initialStatusNotWaiting && !report.doneNotReady && !report.doneNotTeam3L &&
    !report.doneCompletedByNotTeam3L && !report.doneMissingCompletedAt && !report.nonDoneHasCompletion &&
    report.rows === report.expectedRows;
  return report;
}

const UNIT_STATUS_LOG_SHEET = 'STATUS_UNIT_HISTORY';
const UNIT_STATUS_LOG_HEADERS = [
  'UPDATE ID', 'TIMESTAMP', 'USER EMAIL', 'NOMOR POLISI', 'KALIBRASI KE',
  'STATUS SEBELUM', 'STATUS BARU', 'CATATAN', 'SOURCE', 'REASON'
];

const UNIT_STATUS_3L_LOG_SHEET = 'STATUS_3L_HISTORY';
const UNIT_STATUS_3L_LOG_HEADERS = [
  'UPDATE ID', 'TIMESTAMP', 'USER EMAIL', 'NOMOR POLISI', 'KALIBRASI KE',
  'STATUS SEBELUM', 'STATUS BARU', 'CATATAN', 'SOURCE',
  'STATUS UNIT SEBELUM', 'STATUS UNIT SESUDAH', 'REASON'
];
// Batas ini hanya berlaku untuk endpoint Web App. Edit/paste langsung pada
// MASTER_UNIT tidak melewati endpoint ini sehingga tetap dapat memproses
// seluruh baris yang diperlukan oleh admin Sheet.
const STATUS_3L_WEB_BULK_MAX_UNITS = 100;


function getUnitKmHistorySheet_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(UNIT_KM_LOG_SHEET) ||
    spreadsheet.getSheetByName(LEGACY_UNIT_KM_LOG_SHEET);
}


function ensureUnitKmHistorySheet_(ss, report) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  let history = spreadsheet.getSheetByName(UNIT_KM_LOG_SHEET);
  const legacy = spreadsheet.getSheetByName(LEGACY_UNIT_KM_LOG_SHEET);
  if (!history && legacy) {
    legacy.setName(UNIT_KM_LOG_SHEET);
    history = legacy;
    if (report) {
      if (!Array.isArray(report.renamedSheets)) report.renamedSheets = [];
      report.renamedSheets.push(LEGACY_UNIT_KM_LOG_SHEET + ' -> ' + UNIT_KM_LOG_SHEET);
    }
  }
  return history || getOrCreateSheet_(spreadsheet, UNIT_KM_LOG_SHEET);
}


function newCalibrationHistoryId_(timestamp) {
  const date = timestamp instanceof Date ? timestamp : new Date();
  const stamp = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmssSSS');
  return 'HIST-KM-' + stamp + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
}


function setupMasterUnitKmDatabase(sessionToken) {
  const actor = requirePermission_('ADMIN_SETUP', '', { token: sessionToken });
  KM_REQUEST_ACTOR_ = actor;
  return setupMasterUnitKmDatabase_();
}


function setupMasterUnitKmDatabase_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet_(ss, KM_APP.SHEETS.MASTER_UNIT);
    let headers = sheet.getLastColumn() ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0] : [];
    const required = { nopol: 'NOMOR POLISI', operatingPoint: 'OPERATING POINT', toh: 'NAMA TOH/WOH', ...UNIT_KM_FIELDS, ...UNIT_MASTER_DISPLAY_FIELDS };
    let indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
    const added = [];
    // Tolak header ambigu sebelum melakukan perubahan struktur.
    Object.keys(required).forEach(field => {
      const count = headers.filter(header => MASTER_UNIT_FIELD_ALIASES[field].some(alias => normalizeHeaderKey_(header) === normalizeHeaderKey_(alias))).length;
      if (count > 1) throw new Error('Header ganda: ' + required[field] + '. Perbaiki sebelum setup.');
    });
    const existingLog = getUnitKmHistorySheet_(ss);
    if (existingLog) ensureHistoryHeaders_(existingLog, UNIT_KM_LOG_HEADERS);
    const existingStatusLog = ss.getSheetByName(UNIT_STATUS_LOG_SHEET);
    if (existingStatusLog) ensureHistoryHeaders_(existingStatusLog, UNIT_STATUS_LOG_HEADERS);
    const existingStatus3LLog = ss.getSheetByName(UNIT_STATUS_3L_LOG_SHEET);
    if (existingStatus3LLog) ensureHistoryHeaders_(existingStatus3LLog, UNIT_STATUS_3L_LOG_HEADERS);
    const missingFields = Object.keys(required).filter(field => indexes[field] < 0);
    const missingHeaders = missingFields.map(field => required[field]);
    if (missingHeaders.length) {
      const requiredColumns = headers.length + missingHeaders.length;
      if (requiredColumns > sheet.getMaxColumns()) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
      }
      sheet.getRange(1, headers.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
      missingFields.forEach((field, index) => { indexes[field] = headers.length + index; });
      headers.push(...missingHeaders);
      added.push(...missingHeaders);
    }

    // Normalisasi nama header ke format canonical agar tidak tergantung alias lama.
    Object.keys(required).forEach(field => {
      if (indexes[field] >= 0) {
        sheet.getRange(1, indexes[field] + 1).setValue(required[field]);
        headers[indexes[field]] = required[field];
      }
    });

    // Tanggal KM/waktu input ditempatkan lebih dulu, lalu field workflow 3L.
    const ordered = ensureMasterUnitWorkflowColumnOrder_(sheet);
    headers = ordered.headers;
    indexes = ordered.indexes;
    const managementSetup = ensureMasterUnitManagementSchema_(ss, sheet, KM_REQUEST_ACTOR_ || { email: 'SYSTEM_MIGRATION' });
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);

    if (sheet.getMaxRows() < 2) sheet.insertRowsAfter(1, 1);
    // KALIBRASI KE adalah angka hasil sistem dan CATATAN / NOTE adalah teks
    // bebas. Bersihkan validasi pada seluruh kolom, termasuk sel header,
    // sebagai bagian kritis setup agar kegagalan tidak tersamarkan.
    sheet.getRange(1, indexes.calibrationNumber + 1, sheet.getMaxRows(), 1).clearDataValidations();
    sheet.getRange(1, indexes.status3LUpdatedAt + 1, sheet.getMaxRows(), 1).clearDataValidations();
    sheet.getRange(1, indexes.status3LUpdatedBy + 1, sheet.getMaxRows(), 1).clearDataValidations();
    sheet.getRange(1, indexes.status3LNote + 1, sheet.getMaxRows(), 1).clearDataValidations();
    const statusRange = sheet.getRange(2, indexes.statusUnit + 1, sheet.getMaxRows() - 1, 1);
    // Semua business state valid harus berada di dropdown agar hasil trigger tidak
    // berstatus Invalid. Ready to Depart tetap SYSTEM-GENERATED secara business rule;
    // handler onEdit akan menolak pemilihan manual Ready to Depart bila Status 3L != DONE.
    statusRange.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(UNIT_OPERATIONAL_STATUSES, true).setAllowInvalid(false)
      .setHelpText('Kosong diperbolehkan untuk unit baru yang belum masuk workflow. Standby / On Delivery dapat dipilih manual. Ready to Depart otomatis saat Status 3L DONE.')
      .build());
    runNonCritical_('validasi Status 3L MASTER_UNIT', function () {
      applyStatus3LValidation_(sheet, indexes);
    });
    if (sheet.getLastRow() > 1) {
      const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      sheet.getRange(2, indexes.kmPhotoAddedBy + 1, rows.length, 1).setValues(rows.map(row => [
        safeSheetText_(row[indexes.kmPhotoAddedBy] || (row[indexes.kmPhotoId] ? row[indexes.kmUpdatedBy] : ''))
      ]));
    }
    const log = ensureUnitKmHistorySheet_(ss);
    if (!log.getLastRow()) log.getRange(1, 1, 1, UNIT_KM_LOG_HEADERS.length).setValues([UNIT_KM_LOG_HEADERS]);
    const statusLog = getOrCreateSheet_(ss, UNIT_STATUS_LOG_SHEET);
    if (!statusLog.getLastRow()) {
      statusLog.getRange(1, 1, 1, UNIT_STATUS_LOG_HEADERS.length).setValues([UNIT_STATUS_LOG_HEADERS]);
    }
    const status3LLog = getOrCreateSheet_(ss, UNIT_STATUS_3L_LOG_SHEET);
    if (!status3LLog.getLastRow()) {
      status3LLog.getRange(1, 1, 1, UNIT_STATUS_3L_LOG_HEADERS.length).setValues([UNIT_STATUS_3L_LOG_HEADERS]);
    }
    // Migrasi idempotent: histori lama tetap utuh dan hanya field nomor
    // kalibrasi yang kosong yang dilengkapi berdasarkan nopol + urutan waktu.
    backfillCalibrationNumbers_(sheet, indexes, log, status3LLog, statusLog);
    const historicalOpenRows = backfillMissingCalibrationOpenRows_(sheet, indexes, log);
    backfillCalibrationHistoryMetadata_(log);
    const status3LAttributionRepair = backfillStatus3LTimestamps_(sheet, indexes, status3LLog);
    const reconciled = reconcileMasterUnitWorkflowState_(sheet, indexes);
    // v21 safety: never mirror MASTER_UNIT back into audit history as proof.
    // History can repair MASTER_UNIT, but MASTER_UNIT cannot self-validate by
    // creating a SYSTEM_REPAIR event and then reading that event as evidence.
    const summaryHistoryRepair = { written: 0, skippedUnsafeMasterMirroring: true };
    const summaryStatusUnitRepair = { written: 0, skippedUnsafeMasterMirroring: true };
    const historicalSnapshots = backfillCompletedCalibrationSnapshots_(status3LLog);
    // Setelah seluruh evidence status/final snapshot tersedia, lengkapi context
    // HISTORY_KALIBRASI_KM yang kosong secara evidence-based dan idempotent.
    const historyContextRepair = backfillCalibrationHistoryContext_(ss);
    runNonCritical_('format kolom KM MASTER_UNIT', function () {
      formatMasterUnitKmSheet_(sheet, indexes);
    });
    runNonCritical_('format HISTORY_KALIBRASI_KM', function () {
      formatUnitKmHistorySheet_(log);
    });
    runNonCritical_('format STATUS_UNIT_HISTORY', function () {
      formatUnitStatusHistorySheet_(statusLog);
    });
    runNonCritical_('format STATUS_3L_HISTORY', function () {
      formatUnitStatus3LHistorySheet_(status3LLog);
    });
    runNonCritical_('hapus proteksi otomatis lama', removeAutomaticSheetProtections_);
    runNonCritical_('sembunyikan history pengajuan kalibrasi', function () {
      if (!log.isSheetHidden()) log.hideSheet();
    });
    // Trigger installable wajib agar event.user tersedia dan penulisan history
    // tetap memiliki attribution yang dapat diaudit.
    const editTrigger = ensureKalibrasiKmEditTrigger_();
    const status3LWatchTrigger = ensureStatus3LWatchTrigger_();
    SpreadsheetApp.flush();
    invalidateAppCaches_();
    return { success: true, addedColumns: added, historySheet: UNIT_KM_LOG_SHEET,
      statusHistorySheet: UNIT_STATUS_LOG_SHEET,
      status3LHistorySheet: UNIT_STATUS_3L_LOG_SHEET,
      reconciledWorkflowRows: reconciled,
      historicalOpenRows: historicalOpenRows,
      historicalSnapshots: historicalSnapshots,
      historyContextRepair: historyContextRepair,
      summaryHistoryRepair: summaryHistoryRepair,
      summaryStatusUnitRepair: summaryStatusUnitRepair,
      status3LAttributionRepair: status3LAttributionRepair,
      reorderedWorkflowColumns: ordered.reordered,
      historyHidden: log.isSheetHidden(),
      editTrigger: editTrigger,
      status3LWatchTrigger: status3LWatchTrigger,
      masterUnitManagement: managementSetup,
      note: 'HISTORY_KALIBRASI_KM menyimpan pembukaan/final snapshot per siklus. Snapshot bisnis tidak ditimpa; context lifecycle yang kosong dapat dilengkapi dari evidence STATUS_3L_HISTORY, movement history, dan current cycle MASTER_UNIT. Ready to Depart otomatis saat 3L DONE.' };
  } finally { lock.releaseLock(); }
}


function ensureMasterUnitWorkflowColumnOrder_(sheet) {
  let headers = sheet.getLastColumn()
    ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(String)
    : [];
  let indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  const orderedFields = [
    'kmDate', 'kmUpdatedAt', 'kmUpdatedBy',
    'calibrationNumber', 'status3L', 'status3LUpdatedAt', 'status3LUpdatedBy', 'status3LNote'
  ];
  if (orderedFields.some(field => indexes[field] < 0)) {
    return { headers, indexes, reordered: false };
  }

  let reordered = false;
  for (let index = 1; index < orderedFields.length; index++) {
    const previousField = orderedFields[index - 1];
    const field = orderedFields[index];
    if (indexes[previousField] + 1 === indexes[field]) continue;
    // destinationIndex memakai koordinat sebelum pemindahan dan menempatkan
    // kolom tepat sesudah field sebelumnya, baik sumber ada di kiri/kanan.
    sheet.moveColumns(
      sheet.getRange(1, indexes[field] + 1, 1, 1),
      indexes[previousField] + 2
    );
    reordered = true;
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(String);
    indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  }

  const validOrder = orderedFields.every((field, index) =>
    index === 0 || indexes[field] === indexes[orderedFields[index - 1]] + 1
  );
  if (!validOrder) {
    throw new Error('Urutan kolom workflow MASTER_UNIT gagal dinormalisasi. Periksa merge/proteksi kolom lalu ulangi setup.');
  }
  return { headers, indexes, reordered };
}

function formatMasterUnitKmSheet_(sheet, indexes) {
  const dataRows = Math.max(sheet.getMaxRows() - 1, 1);
  ['kmOdometer'].forEach(field => {
    const range = sheet.getRange(2, indexes[field] + 1, dataRows, 1);
    trySetNumberFormat_(range, '#,##0');
    const firstCell = sheet.getRange(2, indexes[field] + 1).getA1Notation();
    range.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied('=OR(ISBLANK(' + firstCell + '),AND(ISNUMBER(' + firstCell + '),' +
        firstCell + '>=0,' + firstCell + '<=999999999,MOD(' + firstCell + ',1)=0))')
      .setAllowInvalid(false)
      .setHelpText('Masukkan digit angka bulat 0-9 tanpa koma, titik, garis miring, atau simbol lain.')
      .build());
    sheet.setColumnWidth(indexes[field] + 1, 135);
  });
  trySetNumberFormat_(sheet.getRange(2, indexes.kmDate + 1, dataRows, 1), 'dd mmmm yyyy');
  ['kmUpdatedAt', 'kmPhotoAddedAt'].forEach(field =>
    trySetNumberFormat_(sheet.getRange(2, indexes[field] + 1, dataRows, 1), 'dd mmmm yyyy, HH.mm.ss')
  );
  trySetNumberFormat_(sheet.getRange(2, indexes.calibrationNumber + 1, dataRows, 1), '0');
  trySetNumberFormat_(sheet.getRange(2, indexes.status3L + 1, dataRows, 1), '@');
  trySetNumberFormat_(sheet.getRange(2, indexes.status3LUpdatedAt + 1, dataRows, 1), 'dd mmmm yyyy, HH.mm.ss');
  trySetNumberFormat_(sheet.getRange(2, indexes.status3LUpdatedBy + 1, dataRows, 1), '@');
  trySetNumberFormat_(sheet.getRange(2, indexes.status3LNote + 1, dataRows, 1), '@');
  sheet.setColumnWidth(indexes.kmDate + 1, 145);
  sheet.setColumnWidth(indexes.calibrationNumber + 1, 115);
  sheet.setColumnWidth(indexes.status3LUpdatedAt + 1, 170);
  sheet.setColumnWidth(indexes.status3LUpdatedBy + 1, 210);
  sheet.setColumnWidth(indexes.status3LNote + 1, 240);
  sheet.setColumnWidth(indexes.kmUpdatedAt + 1, 165);
  sheet.setColumnWidth(indexes.kmUpdatedBy + 1, 210);
}

/**
 * Memindahkan histori lama ke ledger pengajuan baru tanpa menghapus sumber.
 * UPDATE ID lama dipakai sebagai kunci idempotensi; snapshot MASTER_UNIT hanya
 * ditambahkan bila pengajuannya belum ada berdasarkan nopol+kalibrasi/foto.
 */

function formatUnitKmHistorySheet_(sheet) {
  const map = ensureHistoryHeaders_(sheet, UNIT_KM_LOG_HEADERS);
  const dataRows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .setBackground('#16324F').setFontColor('#FFFFFF').setFontWeight('bold');
  trySetNumberFormat_(sheet.getRange(2, map['TIMESTAMP'], dataRows, 1), 'dd mmmm yyyy, HH.mm.ss');
  trySetNumberFormat_(sheet.getRange(2, map['TANGGAL KM'], dataRows, 1), 'dd mmmm yyyy');
  trySetNumberFormat_(sheet.getRange(2, map['FOTO DITAMBAHKAN AT'], dataRows, 1), 'dd mmmm yyyy, HH.mm.ss');
  trySetNumberFormat_(sheet.getRange(2, map['STATUS 3L UPDATED AT'], dataRows, 1), 'dd mmmm yyyy, HH.mm.ss');
  trySetNumberFormat_(sheet.getRange(2, map['COMPLETED AT'], dataRows, 1), 'dd mmmm yyyy, HH.mm.ss');
  trySetNumberFormat_(sheet.getRange(2, map['KM ODOMETER'], dataRows, 1), '#,##0');
  trySetNumberFormat_(sheet.getRange(2, map['KALIBRASI KE'], dataRows, 1), '0');
  trySetNumberFormat_(sheet.getRange(2, map['CALIBRATION SEQUENCE'], dataRows, 1), '0');
  sheet.setColumnWidth(map['OPERATING POINT'], 180);
  sheet.setColumnWidth(map['NAMA TOH/WOH'], 190);
  sheet.setColumnWidth(map['STATUS 3L AWAL'], 150);
  sheet.setColumnWidth(map['STATUS 3L UPDATED AT'], 175);
  sheet.setColumnWidth(map['STATUS 3L UPDATED BY'], 220);
  sheet.setColumnWidth(map['COMPLETED AT'], 175);
  sheet.setColumnWidth(map['COMPLETED BY'], 220);
  sheet.setColumnWidth(map['HISTORY ID'], 230);
  sheet.setColumnWidth(map['SOURCE'], 140);
  sheet.setColumnWidth(map['REASON'], 190);
}

function formatUnitStatusHistorySheet_(sheet) {
  const map = ensureHistoryHeaders_(sheet, UNIT_STATUS_LOG_HEADERS);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .setBackground('#16324F').setFontColor('#FFFFFF').setFontWeight('bold');
  trySetNumberFormat_(
    sheet.getRange(2, map['TIMESTAMP'], Math.max(sheet.getMaxRows() - 1, 1), 1),
    'dd mmmm yyyy, HH.mm.ss'
  );
  trySetNumberFormat_(
    sheet.getRange(2, map['KALIBRASI KE'], Math.max(sheet.getMaxRows() - 1, 1), 1),
    '0'
  );
  sheet.setColumnWidth(map['CATATAN'], 320);
  sheet.setColumnWidth(map['SOURCE'], 150);
  sheet.setColumnWidth(map['REASON'], 190);
}


function formatUnitStatus3LHistorySheet_(sheet) {
  const map = ensureHistoryHeaders_(sheet, UNIT_STATUS_3L_LOG_HEADERS);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .setBackground('#0B5D3A').setFontColor('#FFFFFF').setFontWeight('bold');
  trySetNumberFormat_(
    sheet.getRange(2, map['TIMESTAMP'], Math.max(sheet.getMaxRows() - 1, 1), 1),
    'dd mmmm yyyy, HH.mm.ss'
  );
  trySetNumberFormat_(
    sheet.getRange(2, map['KALIBRASI KE'], Math.max(sheet.getMaxRows() - 1, 1), 1),
    '0'
  );
  sheet.setColumnWidth(map['CATATAN'], 340);
  sheet.setColumnWidth(map['SOURCE'], 160);
  sheet.setColumnWidth(map['REASON'], 190);
}

function getStatus3LConfigRange_() {
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.CONFIG);
  if (!config || config.getLastColumn() < 1) return null;
  const headers = config.getRange(1, 1, 1, config.getLastColumn()).getDisplayValues()[0];
  const columnIndex = headers.findIndex(header => normalizeHeaderKey_(header) === normalizeHeaderKey_('STATUS PROSES 3L'));
  if (columnIndex < 0) return null;
  return config.getRange(2, columnIndex + 1, Math.max(config.getMaxRows() - 1, 1), 1);
}

function applyStatus3LValidation_(masterSheet, indexes) {
  if (!masterSheet || !indexes || indexes.status3L < 0) return;
  const target = masterSheet.getRange(2, indexes.status3L + 1, Math.max(masterSheet.getMaxRows() - 1, 1), 1);
  const source = getStatus3LConfigRange_();
  if (!source) {
    target.clearDataValidations();
    return;
  }
  target.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(source, true)
    .setAllowInvalid(false)
    .setHelpText('Pilih Status 3L dari MASTER_CONFIG > STATUS PROSES 3L.')
    .build());
}

function parseCalibrationNumber_(value) {
  const number = Number(value);
  return Number.isInteger(number) && number > 0 ? number : 0;
}



/**
 * Schema v21 — integrity context HISTORY_KALIBRASI_KM.
 *
 * Tujuan:
 * - melengkapi context coverage dan lifecycle yang kosong tanpa menimpa nilai valid;
 * - menggunakan STATUS_3L_HISTORY sebagai source of truth status/attribution;
 * - menggunakan MASTER_UNIT_MOVEMENT_HISTORY untuk coverage historis bila tersedia;
 * - COMPLETED AT/BY hanya diwajibkan untuk siklus yang memang sudah selesai/locked;
 * - field yang tidak memiliki evidence kuat dibiarkan kosong dan dilaporkan unresolved,
 *   bukan ditebak dari user pengaju odometer.
 */
function calibrationHistoryContextKey_(nopol, calibrationNumber) {
  const plate = normalizeNopol_(nopol);
  const sequence = parseCalibrationNumber_(calibrationNumber);
  return plate && sequence ? plate + '|' + sequence : '';
}

function calibrationHistoryDate_(value) {
  const parsed = value instanceof Date ? value : parseDateTimeInput_(value);
  return parsed instanceof Date && !isNaN(parsed.getTime()) ? parsed : null;
}

function calibrationHistoryActor_(value) {
  return status3LActorDisplay_(displayValue_(value));
}

function calibrationHistoryContextValue_(objects, keys) {
  const sources = Array.isArray(objects) ? objects : [objects];
  for (let i = 0; i < sources.length; i++) {
    const object = sources[i];
    if (!object || typeof object !== 'object') continue;
    for (let j = 0; j < keys.length; j++) {
      const value = object[keys[j]];
      if (cleanText_(value)) return value;
    }
  }
  return '';
}

function buildCalibrationHistoryStatusIndex_(status3LSheet) {
  return buildTrustedStatus3LStateIndex_(status3LSheet);
}

function calibrationHistoryStatusEventForRow_(bucket, rowTimestamp, recordType, lifecycleStatus) {
  if (!bucket) return null;
  const type = normalizeLookupText_(recordType);
  const lifecycle = normalizeLookupText_(lifecycleStatus);
  if (type === 'FINAL_SNAPSHOT' && lifecycle === 'LOCKED') return bucket.done || bucket.latest || bucket.initial;
  if (type === 'CALIBRATION_OPENED') return bucket.initial || bucket.latest;
  const timestamp = calibrationHistoryDate_(rowTimestamp);
  if (timestamp && bucket.events && bucket.events.length) {
    let selected = null;
    bucket.events.forEach(event => {
      if (event.timestamp && event.timestamp.getTime() <= timestamp.getTime()) selected = event;
    });
    if (selected) return selected;
  }
  return bucket.latest || bucket.initial;
}

function buildCalibrationHistoryMovementIndex_(sheet) {
  const result = Object.create(null);
  if (!sheet || sheet.getLastRow() < 2 ||
      !historyHeadersReady_(sheet, MASTER_UNIT_MANAGEMENT.MOVEMENT_HEADERS)) return result;
  const map = getHistoryHeaderMap_(sheet, MASTER_UNIT_MANAGEMENT.MOVEMENT_HEADERS);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  rows.forEach((row, offset) => {
    const timestamp = calibrationHistoryDate_(row[map['TIMESTAMP'] - 1]);
    const event = {
      timestamp: timestamp,
      time: timestamp ? timestamp.getTime() : offset,
      previousOperatingPoint: displayValue_(row[map['PREVIOUS OPERATING POINT'] - 1]),
      newOperatingPoint: displayValue_(row[map['NEW OPERATING POINT'] - 1]),
      previousToh: displayValue_(row[map['PREVIOUS TOH/WOH'] - 1]),
      newToh: displayValue_(row[map['NEW TOH/WOH'] - 1])
    };
    const plates = [
      row[map['NOMOR POLISI'] - 1],
      row[map['PREVIOUS NOMOR POLISI'] - 1],
      row[map['NEW NOMOR POLISI'] - 1]
    ].map(normalizeNopol_).filter(Boolean);
    [...new Set(plates)].forEach(plate => {
      if (!result[plate]) result[plate] = [];
      result[plate].push(event);
    });
  });
  Object.keys(result).forEach(key => result[key].sort((a, b) => a.time - b.time));
  return result;
}

function calibrationHistoryMovementContextAt_(index, nopol, timestamp) {
  const events = index && index[normalizeNopol_(nopol)] || [];
  if (!events.length) return null;
  const target = calibrationHistoryDate_(timestamp);
  const targetTime = target ? target.getTime() : Number.POSITIVE_INFINITY;
  let selected = null;
  for (let i = 0; i < events.length; i++) {
    const event = events[i];
    if (event.time <= targetTime) {
      selected = {
        operatingPoint: event.newOperatingPoint || event.previousOperatingPoint || '',
        toh: event.newToh || event.previousToh || '',
        source: 'MASTER_UNIT_MOVEMENT_HISTORY'
      };
      continue;
    }
    if (!selected && (event.previousOperatingPoint || event.previousToh)) {
      selected = {
        operatingPoint: event.previousOperatingPoint || '',
        toh: event.previousToh || '',
        source: 'MASTER_UNIT_MOVEMENT_HISTORY_PREVIOUS'
      };
    }
    break;
  }
  return selected;
}

function buildCalibrationHistoryMasterIndex_(masterSheet) {
  const result = Object.create(null);
  if (!masterSheet || masterSheet.getLastRow() < 2) return result;
  const headers = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  if (indexes.nopol < 0) return result;
  const rows = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, masterSheet.getLastColumn()).getValues();
  rows.forEach(row => {
    const key = normalizeNopol_(row[indexes.nopol]);
    if (!key) return;
    result[key] = {
      operatingPoint: indexes.operatingPoint >= 0 ? displayValue_(row[indexes.operatingPoint]) : '',
      toh: indexes.toh >= 0 ? displayValue_(row[indexes.toh]) : '',
      calibrationNumber: indexes.calibrationNumber >= 0 ? parseCalibrationNumber_(row[indexes.calibrationNumber]) : 0,
      status3L: indexes.status3L >= 0 ? displayValue_(row[indexes.status3L]) : '',
      status3LUpdatedAt: indexes.status3LUpdatedAt >= 0 ? row[indexes.status3LUpdatedAt] : '',
      status3LUpdatedBy: indexes.status3LUpdatedBy >= 0 ? calibrationHistoryActor_(row[indexes.status3LUpdatedBy]) : ''
    };
  });
  return result;
}

function calibrationHistoryUniqueContext_(sets, key) {
  const values = sets && sets[key] ? [...sets[key]].filter(Boolean) : [];
  return values.length === 1 ? values[0] : '';
}


/**
 * Schema v21 — stronger evidence resolver for legacy HISTORY_KALIBRASI_KM.
 * Tidak membuat timestamp acak. Fallback hanya memakai invariant workflow,
 * sibling context pada siklus yang sama, MASTER_CONFIG mapping unik, atau
 * activity/status ledger yang masih dapat dikaitkan ke siklus yang sama.
 */
function calibrationHistoryUniqueDate_(sets, key) {
  const values = sets && sets[key] ? [...sets[key]].filter(Boolean) : [];
  if (values.length !== 1) return '';
  return calibrationHistoryDate_(values[0]) || values[0];
}

function calibrationHistoryUniqueActor_(sets, key) {
  const values = sets && sets[key] ? [...sets[key]].map(calibrationHistoryActor_).filter(Boolean) : [];
  const unique = [...new Set(values.map(value => normalizeLookupText_(value)))];
  if (unique.length !== 1) return '';
  return values.find(value => normalizeLookupText_(value) === unique[0]) || '';
}

function buildCalibrationHistoryCoverageResolver_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const byOp = Object.create(null);
  const byToh = Object.create(null);
  const add = function (operatingPoint, toh) {
    const op = cleanText_(operatingPoint);
    const tohValue = cleanText_(toh);
    if (op) {
      const key = coverageOperatingPointKey_(op);
      if (!byOp[key]) byOp[key] = { operatingPoints: new Set(), tohs: new Set() };
      byOp[key].operatingPoints.add(op);
      if (tohValue) byOp[key].tohs.add(tohValue);
    }
    if (tohValue) {
      const key = coverageTohKey_(tohValue);
      if (!byToh[key]) byToh[key] = { operatingPoints: new Set(), tohs: new Set() };
      byToh[key].tohs.add(tohValue);
      if (op) byToh[key].operatingPoints.add(op);
    }
  };

  const config = ss.getSheetByName(KM_APP.SHEETS.CONFIG);
  if (config && config.getLastRow() > 1 && config.getLastColumn() > 0) {
    const directory = masterConfigCoverageColumns_(config);
    if (directory && directory.indexes.operatingPoint >= 0) {
      const values = config.getRange(2, 1, config.getLastRow() - 1, config.getLastColumn()).getDisplayValues();
      values.forEach(row => add(
        row[directory.indexes.operatingPoint],
        directory.indexes.toh >= 0 ? row[directory.indexes.toh] : ''
      ));
    }
  }

  const master = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  if (master && master.getLastRow() > 1 && master.getLastColumn() > 0) {
    const headers = master.getRange(1, 1, 1, master.getLastColumn()).getDisplayValues()[0];
    const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
    if (indexes.operatingPoint >= 0) {
      const values = master.getRange(2, 1, master.getLastRow() - 1, master.getLastColumn()).getDisplayValues();
      values.forEach(row => add(
        row[indexes.operatingPoint],
        indexes.toh >= 0 ? row[indexes.toh] : ''
      ));
    }
  }

  return {
    tohForOperatingPoint: function (operatingPoint) {
      const bucket = byOp[coverageOperatingPointKey_(operatingPoint)];
      return bucket && bucket.tohs.size === 1 ? [...bucket.tohs][0] : '';
    },
    operatingPointForToh: function (toh) {
      const bucket = byToh[coverageTohKey_(toh)];
      return bucket && bucket.operatingPoints.size === 1 ? [...bucket.operatingPoints][0] : '';
    }
  };
}

function safeWaitingStatus3LForHistory_() {
  try { return resolveStatus3L_('WAITING'); } catch (error) { return 'WAITING'; }
}

function calibrationHistoryActorFallback_(value) {
  const actor = calibrationHistoryActor_(value);
  const key = normalizeLookupText_(actor);
  if (!actor || key === 'SYSTEM' || key === normalizeLookupText_(KM_STATUS_3L_LEGACY_UNRESOLVED)) {
    return KM_STATUS_3L_SHEET_FALLBACK;
  }
  return actor;
}

function buildCalibrationHistoryContextPlan_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const history = ss.getSheetByName(UNIT_KM_LOG_SHEET);
  const report = {
    schemaReady: false,
    totalRows: 0,
    completedCycleRows: 0,
    changedCells: 0,
    changedRows: 0,
    normalizedLegacyActors: 0,
    unresolvedRows: [],
    missingBefore: {},
    missingAfter: {},
    changesByField: {},
    coverageSources: {},
    evidencePolicy: 'V21_STRONG_EVIDENCE_PLUS_WORKFLOW_INVARIANTS',
    ready: false
  };
  UNIT_KM_HISTORY_CONTEXT_HEADERS.forEach(header => {
    report.missingBefore[header] = 0;
    report.missingAfter[header] = 0;
    report.changesByField[header] = 0;
  });
  if (!history || !historyHeadersReady_(history, UNIT_KM_LOG_HEADERS)) {
    return { history: history, report: report, map: null, rows: [], output: null };
  }
  report.schemaReady = true;
  const map = getHistoryHeaderMap_(history, UNIT_KM_LOG_HEADERS);
  if (history.getLastRow() < 2) {
    report.ready = true;
    return { history: history, report: report, map: map, rows: [], output: {} };
  }
  const rows = history.getRange(2, 1, history.getLastRow() - 1, history.getLastColumn()).getValues();
  report.totalRows = rows.length;
  const statusSheet = ss.getSheetByName(UNIT_STATUS_3L_LOG_SHEET);
  const statusIndex = buildCalibrationHistoryStatusIndex_(statusSheet);
  const auditMap = buildStatus3LAuditEvidenceMap_(statusSheet);
  const movementIndex = buildCalibrationHistoryMovementIndex_(ss.getSheetByName(MASTER_UNIT_MANAGEMENT.MOVEMENT_SHEET));
  const masterIndex = buildCalibrationHistoryMasterIndex_(ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT));
  const activityIndex = buildStatus3LActivityEvidenceMap_(ss.getSheetByName(KM_APP.SHEETS.LOG));
  const coverageResolver = buildCalibrationHistoryCoverageResolver_(ss);

  // Kumpulkan context existing per siklus. Hanya nilai unik yang boleh dipropagasikan.
  const cycleContextSets = Object.create(null);
  const cycleCompletion = Object.create(null);
  rows.forEach((row, offset) => {
    const key = calibrationHistoryContextKey_(row[map['NOMOR POLISI'] - 1], row[map['KALIBRASI KE'] - 1]);
    if (!key) return;
    if (!cycleContextSets[key]) cycleContextSets[key] = {
      operatingPoint: new Set(), toh: new Set(), status3LInitial: new Set(),
      status3LUpdatedAt: new Set(), status3LUpdatedBy: new Set(),
      completedAt: new Set(), completedBy: new Set()
    };
    const bucket = cycleContextSets[key];
    const op = cleanText_(row[map['OPERATING POINT'] - 1]);
    const toh = cleanText_(row[map['NAMA TOH/WOH'] - 1]);
    const initial = cleanText_(row[map['STATUS 3L AWAL'] - 1]);
    if (op) bucket.operatingPoint.add(op);
    if (toh) bucket.toh.add(toh);
    if (initial) bucket.status3LInitial.add(initial);
    const statusAtExisting = calibrationHistoryDate_(row[map['STATUS 3L UPDATED AT'] - 1]);
    const statusByExisting = calibrationHistoryActor_(row[map['STATUS 3L UPDATED BY'] - 1]);
    const completedAtExisting = calibrationHistoryDate_(row[map['COMPLETED AT'] - 1]);
    const completedByExisting = calibrationHistoryActor_(row[map['COMPLETED BY'] - 1]);
    if (statusAtExisting) bucket.status3LUpdatedAt.add(statusAtExisting.toISOString());
    if (statusByExisting) bucket.status3LUpdatedBy.add(statusByExisting);
    if (completedAtExisting) bucket.completedAt.add(completedAtExisting.toISOString());
    if (completedByExisting) bucket.completedBy.add(completedByExisting);

    const recordType = normalizeLookupText_(row[map['RECORD TYPE'] - 1]);
    const lifecycle = normalizeLookupText_(row[map['LIFECYCLE STATUS'] - 1]);
    const completedAt = calibrationHistoryDate_(row[map['COMPLETED AT'] - 1]);
    const completedBy = calibrationHistoryActor_(row[map['COMPLETED BY'] - 1]);
    if (recordType === 'FINAL_SNAPSHOT' && lifecycle === 'LOCKED') {
      const trustedDone = statusIndex[key] && statusIndex[key].done;
      if (!trustedDone && !isDirectTrustedFinalSnapshot_(row, map)) return;
      const timestamp = completedAt || calibrationHistoryDate_(row[map['STATUS 3L UPDATED AT'] - 1]) ||
        calibrationHistoryDate_(row[map['TIMESTAMP'] - 1]);
      const actor = completedBy || calibrationHistoryActor_(row[map['STATUS 3L UPDATED BY'] - 1]) ||
        calibrationHistoryActor_(row[map['USER EMAIL'] - 1]);
      const current = cycleCompletion[key];
      if (!current || (timestamp && (!current.timestamp || timestamp.getTime() < current.timestamp.getTime()))) {
        cycleCompletion[key] = { timestamp: timestamp, updatedBy: actor, source: 'TRUSTED_HISTORY_FINAL_SNAPSHOT' };
      } else if (current && !current.updatedBy && actor) {
        current.updatedBy = actor;
      }
    }
  });

  // STATUS_3L_HISTORY merupakan source of truth completion; override timestamp/actor snapshot bila ada DONE evidence.
  Object.keys(statusIndex).forEach(key => {
    const statusBucket = statusIndex[key];
    if (!statusBucket.done) return;
    const done = statusBucket.done;
    const auditParts = key.split('|');
    const sequence = Number(auditParts[auditParts.length - 1]) || 0;
    const nopolKey = auditParts.slice(0, -1).join('|');
    const directAudit = status3LAuditEvidenceFor_(auditMap, nopolKey, sequence, 'DONE');
    cycleCompletion[key] = {
      timestamp: directAudit && directAudit.timestamp || done.timestamp || (cycleCompletion[key] && cycleCompletion[key].timestamp) || null,
      updatedBy: directAudit && directAudit.updatedBy || done.updatedBy || (cycleCompletion[key] && cycleCompletion[key].updatedBy) || '',
      source: directAudit ? 'STATUS_3L_AUDIT_EVIDENCE' : 'STATUS_3L_HISTORY'
    };
  });

  const output = {};
  UNIT_KM_HISTORY_CONTEXT_HEADERS.forEach(header => {
    output[header] = rows.map(row => [row[map[header] - 1]]);
  });
  const changedRowSet = new Set();

  const assignIfBlank = (header, offset, value, source) => {
    const existing = output[header][offset][0];
    if (cleanText_(existing) || value === '' || value == null) return false;
    output[header][offset][0] = value;
    report.changedCells++;
    report.changesByField[header]++;
    if (source) report.coverageSources[source] = (report.coverageSources[source] || 0) + 1;
    changedRowSet.add(offset + 2);
    return true;
  };

  rows.forEach((row, offset) => {
    const nopol = displayValue_(row[map['NOMOR POLISI'] - 1]);
    const sequence = parseCalibrationNumber_(row[map['KALIBRASI KE'] - 1]);
    const key = calibrationHistoryContextKey_(nopol, sequence);
    if (!key) return;
    const recordType = displayValue_(row[map['RECORD TYPE'] - 1]);
    const lifecycle = displayValue_(row[map['LIFECYCLE STATUS'] - 1]);
    const rowTimestamp = row[map['TIMESTAMP'] - 1];
    const dataBefore = safeParseHistoryJson_(row[map['DATA SEBELUM'] - 1]);
    const dataAfter = safeParseHistoryJson_(row[map['DATA SETELAH'] - 1]);
    const cycleSets = cycleContextSets[key] || {};
    const plateKey = normalizeNopol_(nopol);
    const master = masterIndex[plateKey] || null;

    let operatingPoint = calibrationHistoryContextValue_([dataAfter, dataBefore],
      ['operatingPoint','operating_point','OPERATING POINT']);
    let toh = calibrationHistoryContextValue_([dataAfter, dataBefore],
      ['toh','tohWoh','namaTohWoh','NAMA TOH/WOH','TOH/WOH']);
    let coverageSource = operatingPoint || toh ? 'HISTORY_JSON' : '';
    if (!operatingPoint) operatingPoint = calibrationHistoryUniqueContext_(cycleSets, 'operatingPoint');
    if (!toh) toh = calibrationHistoryUniqueContext_(cycleSets, 'toh');
    if (!coverageSource && (operatingPoint || toh)) coverageSource = 'SAME_CALIBRATION_UNIQUE';

    if (!operatingPoint || !toh) {
      const movement = calibrationHistoryMovementContextAt_(movementIndex, nopol, rowTimestamp);
      if (movement) {
        if (!operatingPoint) operatingPoint = movement.operatingPoint;
        if (!toh) toh = movement.toh;
        if (operatingPoint || toh) coverageSource = movement.source;
      }
    }
    // Current MASTER_UNIT hanya aman sebagai fallback untuk siklus aktif yang sama.
    if (master && master.calibrationNumber === sequence) {
      if (!operatingPoint) operatingPoint = master.operatingPoint;
      if (!toh) toh = master.toh;
      if (!coverageSource && (operatingPoint || toh)) coverageSource = 'MASTER_UNIT_CURRENT_CYCLE';
    }
    // Bila salah satu sisi mapping masih kosong, gunakan MASTER_CONFIG hanya jika
    // pasangan OP↔TOH unik. Ini bukan asumsi dan aman terhadap OP yang punya banyak TOH.
    if (operatingPoint && !toh) {
      const mappedToh = coverageResolver.tohForOperatingPoint(operatingPoint);
      if (mappedToh) { toh = mappedToh; coverageSource = 'MASTER_CONFIG_UNIQUE_MAPPING'; }
    }
    if (toh && !operatingPoint) {
      const mappedOp = coverageResolver.operatingPointForToh(toh);
      if (mappedOp) { operatingPoint = mappedOp; coverageSource = 'MASTER_CONFIG_UNIQUE_MAPPING'; }
    }
    assignIfBlank('OPERATING POINT', offset, safeSheetText_(operatingPoint), coverageSource);
    assignIfBlank('NAMA TOH/WOH', offset, safeSheetText_(toh), coverageSource);

    const statusBucket = statusIndex[key] || null;
    const statusEvent = calibrationHistoryStatusEventForRow_(statusBucket, rowTimestamp, recordType, lifecycle);
    let initialStatus = calibrationHistoryContextValue_([dataAfter, dataBefore], ['status3LInitial','status3LAwal']);
    if (!initialStatus) initialStatus = calibrationHistoryUniqueContext_(cycleSets, 'status3LInitial');
    if (!initialStatus && statusBucket && statusBucket.initial) initialStatus = statusBucket.initial.status3L;
    // CALIBRATION_OPENED dibentuk oleh workflow dengan Status 3L awal WAITING.
    // Untuk final snapshot, invariant hanya dipakai bila pasangan FINAL_SNAPSHOT+LOCKED
    // valid; lifecycle/type parsial tidak boleh dianggap evidence completion.
    if (!initialStatus && (normalizeLookupText_(recordType) === 'CALIBRATION_OPENED' ||
        (normalizeLookupText_(recordType) === 'FINAL_SNAPSHOT' && normalizeLookupText_(lifecycle) === 'LOCKED'))) {
      initialStatus = safeWaitingStatus3LForHistory_();
    }
    assignIfBlank('STATUS 3L AWAL', offset, safeSheetText_(initialStatus),
      statusBucket && statusBucket.initial ? 'STATUS_3L_HISTORY' : 'WORKFLOW_INVARIANT_WAITING');

    let statusUpdatedAt = statusEvent && statusEvent.timestamp ? statusEvent.timestamp : '';
    let statusUpdatedBy = statusEvent && statusEvent.updatedBy ? statusEvent.updatedBy : '';
    if (!statusUpdatedAt) statusUpdatedAt = calibrationHistoryUniqueDate_(cycleSets, 'status3LUpdatedAt');
    if (!statusUpdatedBy) statusUpdatedBy = calibrationHistoryUniqueActor_(cycleSets, 'status3LUpdatedBy');
    if (!statusUpdatedAt) statusUpdatedAt = calibrationHistoryContextValue_([dataAfter, dataBefore], ['status3LUpdatedAt']);
    if (!statusUpdatedBy) statusUpdatedBy = calibrationHistoryActor_(
      calibrationHistoryContextValue_([dataAfter, dataBefore], ['status3LUpdatedBy'])
    );
    // Secondary legacy evidence dari KM_ACTIVITY_LOG, dibatasi mulai TIMESTAMP row/siklus.
    if ((!statusUpdatedAt || !statusUpdatedBy) && statusEvent && cleanText_(statusEvent.status3L)) {
      const activityEvidence = status3LActivityEvidenceFor_(
        activityIndex, nopol, statusEvent.status3L, rowTimestamp, sequence
      );
      if (activityEvidence) {
        if (!statusUpdatedAt) statusUpdatedAt = activityEvidence.timestamp || activityEvidence.updatedAt;
        if (!statusUpdatedBy) statusUpdatedBy = activityEvidence.updatedBy;
      }
    }
    // CALIBRATION_OPENED adalah auto transition ke WAITING pada timestamp pembukaan.
    if (normalizeLookupText_(recordType) === 'CALIBRATION_OPENED') {
      if (!statusUpdatedAt) statusUpdatedAt = calibrationHistoryDate_(rowTimestamp) || rowTimestamp;
      if (!statusUpdatedBy) statusUpdatedBy = 'SYSTEM';
    }
    if (master && master.calibrationNumber === sequence && !statusUpdatedAt && master.status3LUpdatedAt) {
      statusUpdatedAt = calibrationHistoryDate_(master.status3LUpdatedAt) || master.status3LUpdatedAt;
    }
    if (master && master.calibrationNumber === sequence && !statusUpdatedBy && master.status3LUpdatedBy) {
      statusUpdatedBy = master.status3LUpdatedBy;
    }
    assignIfBlank('STATUS 3L UPDATED AT', offset,
      calibrationHistoryDate_(statusUpdatedAt) || statusUpdatedAt || '', 'STATUS_3L_HISTORY');
    assignIfBlank('STATUS 3L UPDATED BY', offset, safeSheetText_(statusUpdatedBy), 'STATUS_3L_HISTORY');

    // Normalisasi marker legacy walaupun sel tidak kosong.
    ['STATUS 3L UPDATED BY','COMPLETED BY'].forEach(header => {
      const current = displayValue_(output[header][offset][0]);
      if (normalizeLookupText_(current) === normalizeLookupText_(KM_STATUS_3L_LEGACY_UNRESOLVED)) {
        output[header][offset][0] = KM_STATUS_3L_SHEET_FALLBACK;
        report.changedCells++;
        report.changesByField[header]++;
        report.normalizedLegacyActors++;
        changedRowSet.add(offset + 2);
      }
    });

    let completion = cycleCompletion[key] || null;
    const finalLocked = normalizeLookupText_(recordType) === 'FINAL_SNAPSHOT' && normalizeLookupText_(lifecycle) === 'LOCKED';
    const trustedFinalLocked = finalLocked && Boolean(
      (statusIndex[key] && statusIndex[key].done) || isDirectTrustedFinalSnapshot_(row, map)
    );
    const completedRequired = Boolean(completion) || trustedFinalLocked;
    if (completedRequired) report.completedCycleRows++;
    let completionAt = completion && completion.timestamp || calibrationHistoryUniqueDate_(cycleSets, 'completedAt');
    let completionBy = completion && completion.updatedBy || calibrationHistoryUniqueActor_(cycleSets, 'completedBy');
    // FINAL_SNAPSHOT/LOCKED adalah direct evidence bahwa baris dibuat saat DONE.
    // Writer aplikasi menggunakan TIMESTAMP final sebagai completion time.
    if (completedRequired && !completionAt && trustedFinalLocked) {
      completionAt = calibrationHistoryDate_(rowTimestamp) || rowTimestamp;
    }
    if (completedRequired && !completionBy) {
      completionBy = statusUpdatedBy || calibrationHistoryActor_(row[map['USER EMAIL'] - 1]);
    }
    if (completedRequired && !completionBy) completionBy = KM_STATUS_3L_SHEET_FALLBACK;
    if (completedRequired) {
      const source = completion && completion.source || 'FINAL_SNAPSHOT_DIRECT_EVIDENCE';
      assignIfBlank('COMPLETED AT', offset, completionAt || '', source);
      assignIfBlank('COMPLETED BY', offset, safeSheetText_(calibrationHistoryActorFallback_(completionBy)), source);
      // Final/locked row status update must align with completion when old ledger is incomplete.
      assignIfBlank('STATUS 3L UPDATED AT', offset, completionAt || '', source);
      assignIfBlank('STATUS 3L UPDATED BY', offset, safeSheetText_(calibrationHistoryActorFallback_(completionBy)), source);
    }

    // Hitung missing sesuai semantics setelah plan diterapkan.
    UNIT_KM_HISTORY_CONTEXT_HEADERS.forEach(header => {
      const before = row[map[header] - 1];
      if (!cleanText_(before)) report.missingBefore[header]++;
      const value = output[header][offset][0];
      const required = header === 'COMPLETED AT' || header === 'COMPLETED BY' ? completedRequired : true;
      if (required && !cleanText_(value)) report.missingAfter[header]++;
    });
    const unresolved = UNIT_KM_HISTORY_CONTEXT_HEADERS.filter(header => {
      const required = header === 'COMPLETED AT' || header === 'COMPLETED BY' ? completedRequired : true;
      return required && !cleanText_(output[header][offset][0]);
    });
    if (unresolved.length && report.unresolvedRows.length < 100) {
      report.unresolvedRows.push({
        row: offset + 2,
        nopol: nopol,
        calibrationNumber: sequence,
        recordType: recordType,
        lifecycleStatus: lifecycle,
        missing: unresolved
      });
    }
  });

  report.changedRows = changedRowSet.size;
  report.ready = report.schemaReady && UNIT_KM_HISTORY_CONTEXT_HEADERS.every(header => report.missingAfter[header] === 0);
  return { history: history, report: report, map: map, rows: rows, output: output };
}

function auditCalibrationHistoryContextIntegrity_() {
  return buildCalibrationHistoryContextPlan_(SpreadsheetApp.getActiveSpreadsheet()).report;
}

function applyCalibrationHistoryContextPlan_(plan) {
  if (!plan || !plan.history || !plan.map || !plan.output || !plan.report.changedCells) return plan ? plan.report : {};
  const sheet = plan.history;
  const rowCount = plan.rows.length;
  UNIT_KM_HISTORY_CONTEXT_HEADERS.forEach(header => {
    if (!plan.report.changesByField[header]) return;
    sheet.getRange(2, plan.map[header], rowCount, 1).setValues(plan.output[header]);
  });
  SpreadsheetApp.flush();
  return plan.report;
}

function backfillCalibrationHistoryContext_(spreadsheet) {
  const plan = buildCalibrationHistoryContextPlan_(spreadsheet || SpreadsheetApp.getActiveSpreadsheet());
  return applyCalibrationHistoryContextPlan_(plan);
}

function repairCalibrationHistoryContextIntegrity_(actor, options) {
  const settings = options || {};
  const lock = acquireScriptLock_(30000, 'repair HISTORY_KALIBRASI_KM context');
  try {
    const backup = settings.skipBackup ? null : createAdministrationBackup_(actor, 'BEFORE_HISTORY_CONTEXT_REPAIR');
    const report = backfillCalibrationHistoryContext_(SpreadsheetApp.getActiveSpreadsheet());
    formatUnitKmHistorySheet_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNIT_KM_LOG_SHEET));
    invalidateAppCaches_();
    return { success: true, backup: backup, report: report };
  } finally {
    lock.releaseLock();
  }
}

function repairCalibrationHistoryContext(sessionToken) {
  const actor = requirePermission_('ADMIN_SETUP', '', { token: sessionToken });
  return repairCalibrationHistoryContextIntegrity_(actor);
}

function repairCalibrationHistoryContextFromSheet() {
  const actor = requirePermission_('ADMIN_SETUP', '', { source: 'SHEET' });
  return repairCalibrationHistoryContextIntegrity_(actor);
}

function backfillCalibrationHistoryMetadata_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const map = ensureHistoryHeaders_(sheet, UNIT_KM_LOG_HEADERS);
  const required = ['HISTORY ID', 'REQUEST ID', 'CALIBRATION SEQUENCE', 'RECORD TYPE',
    'LIFECYCLE STATUS', 'SOURCE', 'REASON', 'DATA SETELAH'];
  const rowCount = sheet.getLastRow() - 1;
  const rows = sheet.getRange(2, 1, rowCount, sheet.getLastColumn()).getValues();
  const output = {};
  const changed = {};
  required.forEach(header => {
    output[header] = rows.map(row => [row[map[header] - 1]]);
    changed[header] = 0;
  });
  rows.forEach((row, offset) => {
    if (!row.some(value => value !== '' && value != null)) return;
    const timestamp = row[map['TIMESTAMP'] - 1] instanceof Date
      ? row[map['TIMESTAMP'] - 1] : new Date();
    const updateId = cleanText_(row[map['UPDATE ID'] - 1]) || Utilities.getUuid();
    const put = (header, value) => {
      if (cleanText_(output[header][offset][0])) return;
      output[header][offset][0] = value;
      changed[header]++;
    };
    put('HISTORY ID', newCalibrationHistoryId_(timestamp));
    put('REQUEST ID', updateId);
    put('CALIBRATION SEQUENCE', parseCalibrationNumber_(row[map['KALIBRASI KE'] - 1]) || '');
    put('RECORD TYPE', 'CALIBRATION_OPENED');
    put('LIFECYCLE STATUS', 'ACTIVE');
    put('SOURCE', 'SYSTEM');
    put('REASON', 'LEGACY_MIGRATION');
  });
  let updated = 0;
  required.forEach(header => {
    if (!changed[header]) return;
    sheet.getRange(2, map[header], rowCount, 1).setValues(output[header]);
    updated += changed[header];
  });
  return updated;
}


function backfillMissingCalibrationOpenRows_(masterSheet, indexes, historySheet) {
  if (!masterSheet || masterSheet.getLastRow() < 2 || !historySheet) return 0;
  const map = ensureHistoryHeaders_(historySheet, UNIT_KM_LOG_HEADERS);
  const existingKeys = new Set();
  if (historySheet.getLastRow() > 1) {
    historySheet.getRange(2, 1, historySheet.getLastRow() - 1, historySheet.getLastColumn()).getValues()
      .forEach(row => {
        const type = map['RECORD TYPE'] ? normalizeLookupText_(row[map['RECORD TYPE'] - 1]) : '';
        if (type === 'FINAL_SNAPSHOT') return;
        const key = normalizeNopol_(row[map['NOMOR POLISI'] - 1]) + '|' +
          parseCalibrationNumber_(row[map['KALIBRASI KE'] - 1]);
        if (key !== '|0') existingKeys.add(key);
      });
  }
  const rows = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, masterSheet.getLastColumn()).getValues();
  const additions = [];
  rows.forEach(row => {
    const snapshot = unitKmSnapshot_(row, indexes);
    const nopol = normalizeDisplayNopol_(row[indexes.nopol]);
    const sequence = parseCalibrationNumber_(snapshot.calibrationNumber);
    const key = normalizeNopol_(nopol) + '|' + sequence;
    if (!nopol || !sequence || existingKeys.has(key) || parseUnitKm_(snapshot.kmOdometer) === null ||
        !snapshot.kmDate || !snapshot.kmPhotoId) return;
    const timestamp = row[indexes.kmUpdatedAt] instanceof Date ? row[indexes.kmUpdatedAt] : new Date();
    const email = displayValue_(row[indexes.kmUpdatedBy]) || 'SYSTEM';
    const updateId = Utilities.getUuid();
    additions.push(buildHistoryRow_(historySheet, map, {
      'UPDATE ID': updateId,
      'HISTORY ID': newCalibrationHistoryId_(timestamp),
      'REQUEST ID': updateId,
      'CALIBRATION SEQUENCE': sequence,
      'RECORD TYPE': 'CALIBRATION_OPENED',
      'LIFECYCLE STATUS': 'ACTIVE',
      'TIMESTAMP': timestamp,
      'USER EMAIL': safeSheetText_(email),
      'NOMOR POLISI': safeSheetText_(nopol),
      'KALIBRASI KE': sequence,
      'TANGGAL KM': row[indexes.kmDate],
      'STATUS UNIT': safeSheetText_(row[indexes.statusUnit]),
      'KM ODOMETER': snapshot.kmOdometer,
      'CATATAN': safeSheetText_(row[indexes.status3LNote]),
      'FOTO ODOMETER ID': snapshot.kmPhotoId,
      'FOTO DITAMBAHKAN AT': row[indexes.kmPhotoAddedAt],
      'FOTO DIUNGGAH OLEH': safeSheetText_(row[indexes.kmPhotoAddedBy] || email),
      'OPERATING POINT': safeSheetText_(row[indexes.operatingPoint]),
      'NAMA TOH/WOH': safeSheetText_(row[indexes.toh]),
      'STATUS 3L AWAL': safeSheetText_(row[indexes.status3L]),
      'STATUS 3L UPDATED AT': row[indexes.status3LUpdatedAt],
      'STATUS 3L UPDATED BY': safeSheetText_(row[indexes.status3LUpdatedBy]),
      'SOURCE': 'SYSTEM',
      'REASON': 'LEGACY_MASTER_SNAPSHOT'
    }));
    existingKeys.add(key);
  });
  if (!additions.length) return 0;
  ensureSheetCapacity_(historySheet, additions.length);
  historySheet.getRange(historySheet.getLastRow() + 1, 1, additions.length, historySheet.getLastColumn())
    .setValues(additions);
  return additions.length;
}

/**
 * Melengkapi nomor kalibrasi untuk data legacy tanpa menghapus atau menulis
 * ulang nilai histori lain. Urutan row histori adalah tie-breaker apabila
 * beberapa kalibrasi dilakukan pada tanggal yang sama.
 */
function backfillCalibrationNumbers_(masterSheet, indexes, kmLog, status3LLog, statusLog) {
  const kmMap = ensureHistoryHeaders_(kmLog, UNIT_KM_LOG_HEADERS);
  const status3LMap = ensureHistoryHeaders_(status3LLog, UNIT_STATUS_3L_LOG_HEADERS);
  const statusMap = statusLog ? ensureHistoryHeaders_(statusLog, UNIT_STATUS_LOG_HEADERS) : null;
  const currentByNopol = Object.create(null);
  const timelineByNopol = Object.create(null);

  if (kmLog.getLastRow() > 1) {
    const rowCount = kmLog.getLastRow() - 1;
    const rows = kmLog.getRange(2, 1, rowCount, kmLog.getLastColumn()).getValues();
    const sequenceValues = rows.map((row, offset) => {
      const key = normalizeNopol_(row[kmMap['NOMOR POLISI'] - 1]);
      if (!key) return [''];
      const existing = parseCalibrationNumber_(row[kmMap['KALIBRASI KE'] - 1]);
      const sequence = existing || ((currentByNopol[key] || 0) + 1);
      currentByNopol[key] = Math.max(currentByNopol[key] || 0, sequence);
      if (!timelineByNopol[key]) timelineByNopol[key] = [];
      const timestamp = row[kmMap['TIMESTAMP'] - 1];
      const parsedTimestamp = timestamp instanceof Date ? timestamp : parseDateTimeInput_(timestamp);
      timelineByNopol[key].push({
        sequence: sequence,
        time: parsedTimestamp instanceof Date && !isNaN(parsedTimestamp.getTime())
          ? parsedTimestamp.getTime() : offset
      });
      return [existing ? row[kmMap['KALIBRASI KE'] - 1] : sequence];
    });
    kmLog.getRange(2, kmMap['KALIBRASI KE'], rowCount, 1).setValues(sequenceValues);
  }

  const masterSequenceByNopol = Object.create(null);
  if (masterSheet.getLastRow() > 1 && indexes.calibrationNumber >= 0) {
    const rowCount = masterSheet.getLastRow() - 1;
    const rows = masterSheet.getRange(2, 1, rowCount, masterSheet.getLastColumn()).getValues();
    const sequenceValues = rows.map(row => {
      const key = normalizeNopol_(row[indexes.nopol]);
      if (!key) return [''];
      const existing = parseCalibrationNumber_(row[indexes.calibrationNumber]);
      const hasKm = indexes.kmOdometer >= 0 && parseUnitKm_(row[indexes.kmOdometer]) !== null &&
        indexes.kmDate >= 0 && Boolean(unitKmDate_(row[indexes.kmDate]));
      const sequence = Math.max(existing, currentByNopol[key] || 0, hasKm ? 1 : 0);
      masterSequenceByNopol[key] = sequence;
      return [sequence || ''];
    });
    masterSheet.getRange(2, indexes.calibrationNumber + 1, rowCount, 1).setValues(sequenceValues);
  }

  if (status3LLog.getLastRow() > 1) {
    const rowCount = status3LLog.getLastRow() - 1;
    const rows = status3LLog.getRange(2, 1, rowCount, status3LLog.getLastColumn()).getValues();
    const sequenceValues = rows.map((row, offset) => {
      const existing = parseCalibrationNumber_(row[status3LMap['KALIBRASI KE'] - 1]);
      if (existing) return [row[status3LMap['KALIBRASI KE'] - 1]];
      const key = normalizeNopol_(row[status3LMap['NOMOR POLISI'] - 1]);
      if (!key) return [''];
      const timestamp = row[status3LMap['TIMESTAMP'] - 1];
      const parsedTimestamp = timestamp instanceof Date ? timestamp : parseDateTimeInput_(timestamp);
      const eventTime = parsedTimestamp instanceof Date && !isNaN(parsedTimestamp.getTime())
        ? parsedTimestamp.getTime() : offset;
      const timeline = timelineByNopol[key] || [];
      let sequence = 0;
      timeline.forEach(event => {
        if (event.time <= eventTime) sequence = event.sequence;
      });
      return [sequence || (timeline.length ? timeline[0].sequence : masterSequenceByNopol[key]) || ''];
    });
    status3LLog.getRange(2, status3LMap['KALIBRASI KE'], rowCount, 1).setValues(sequenceValues);
  }

  if (statusLog && statusMap && statusLog.getLastRow() > 1) {
    const rowCount = statusLog.getLastRow() - 1;
    const rows = statusLog.getRange(2, 1, rowCount, statusLog.getLastColumn()).getValues();
    const sequenceValues = rows.map((row, offset) => {
      const existing = parseCalibrationNumber_(row[statusMap['KALIBRASI KE'] - 1]);
      if (existing) return [row[statusMap['KALIBRASI KE'] - 1]];
      const key = normalizeNopol_(row[statusMap['NOMOR POLISI'] - 1]);
      if (!key) return [''];
      const timestamp = row[statusMap['TIMESTAMP'] - 1];
      const parsedTimestamp = timestamp instanceof Date ? timestamp : parseDateTimeInput_(timestamp);
      const eventTime = parsedTimestamp instanceof Date && !isNaN(parsedTimestamp.getTime())
        ? parsedTimestamp.getTime() : offset;
      const timeline = timelineByNopol[key] || [];
      let sequence = 0;
      timeline.forEach(event => {
        if (event.time <= eventTime) sequence = event.sequence;
      });
      return [sequence || (timeline.length ? timeline[0].sequence : masterSequenceByNopol[key]) || ''];
    });
    statusLog.getRange(2, statusMap['KALIBRASI KE'], rowCount, 1).setValues(sequenceValues);
  }
}

function status3LAuditKey_(nopol, calibrationNumber, status3L) {
  const plate = normalizeNopol_(nopol);
  if (!plate) return '';
  const sequence = parseCalibrationNumber_(calibrationNumber) || 0;
  const statusKey = normalizeLookupText_(status3L) || '__EMPTY__';
  return plate + '|' + sequence + '|' + statusKey;
}

function isResolvedStatus3LActor_(value) {
  const actor = status3LActorDisplay_(value);
  const key = normalizeLookupText_(actor);
  if (key === normalizeLookupText_(KM_STATUS_3L_SHEET_FALLBACK) ||
      key === normalizeLookupText_(KM_STATUS_3L_AUTOMATION_ACTOR)) return true;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(actor);
}

const KM_STATUS_3L_WEAK_REPAIR_REASONS = Object.freeze([
  'STATUS_3L_MASTER_STATE_RECONCILED_FROM_HISTORY',
  'STATUS_3L_MASTER_STATE_RECONCILED_FROM_TRUSTED_HISTORY',
  'MASTER_STATE_RECONCILIATION',
  'STATUS_3L_DONE_INVARIANT_REPAIR',
  'STATUS_3L_DONE_INVARIANT',
  'FINAL_SNAPSHOT_INVARIANT',
  'LEGACY_FINAL_SNAPSHOT',
  'REPAIR_FINAL_SNAPSHOT',
  'WORKFLOW_RECONCILE',
  'CURRENT_STATE_RECONCILIATION',
  'CURRENT_STATE',
  'CALIBRATION_DONE_PRESERVED'
]);

const KM_STATUS_3L_TRUSTED_CORRECTION_REASONS = Object.freeze([
  'STATUS_3L_FALSE_COMPLETION_INVALIDATED_FROM_TRUSTED_WAITING',
  'STATUS_3L_ROLLBACK_RECOVERED_FROM_AUDIT',
  'STATUS_3L_DONE_ATTRIBUTION_NORMALIZED_TO_TEAM_3L'
]);

function isWeakStatus3LRepairReason_(reason) {
  const key = normalizeLookupText_(reason);
  return KM_STATUS_3L_WEAK_REPAIR_REASONS.some(item => normalizeLookupText_(item) === key);
}

function isTrustedStatus3LCorrectionReason_(reason) {
  const key = normalizeLookupText_(reason);
  return KM_STATUS_3L_TRUSTED_CORRECTION_REASONS.some(item => normalizeLookupText_(item) === key);
}

/**
 * Classify independent Status 3L evidence. The MASTER_UNIT snapshot and old
 * SYSTEM_REPAIR mirror events are deliberately excluded to prevent circular
 * self-validation (MASTER -> history -> MASTER).
 */
function isTrustedStatus3LEvent_(updatedBy, source, reason, status3L) {
  const actor = status3LActorDisplay_(updatedBy);
  const actorKey = normalizeLookupText_(actor);
  const sourceKey = normalizeLookupText_(source);
  const reasonKey = normalizeLookupText_(reason);
  const statusKey = normalizeLookupText_(status3L);
  if (/NOTE_UPDATED/.test(reasonKey)) return false;
  if (isTrustedStatus3LCorrectionReason_(reasonKey)) return true;
  if (isWeakStatus3LRepairReason_(reasonKey)) return false;
  if (sourceKey === 'SYSTEM_REPAIR') return false;

  if (sourceKey === 'DASHBOARD' || sourceKey === 'SPREADSHEET' || sourceKey === 'SPREADSHEET_PRIMARY') {
    return isResolvedStatus3LActor_(actor) || actorKey === normalizeLookupText_(KM_STATUS_3L_SHEET_FALLBACK);
  }
  if (sourceKey === 'AUTO_TRIGGER') {
    return ['CALIBRATION_CREATED', 'UNIT_ON_DELIVERY', 'UNIT_STANDBY', 'STATUS_3L_DONE', 'STATUS_3L_UPDATED']
      .includes(reasonKey) || (statusKey !== 'DONE' && Boolean(statusKey));
  }
  if (sourceKey === 'EXTERNAL_SYNC') {
    // Perubahan automation/API atau external write tidak memicu onEdit. Event ini dibuat watcher
    // setelah divergence MASTER_UNIT vs ledger tepercaya tetap ada.
    return isResolvedStatus3LActor_(actor) &&
      (Boolean(statusKey) || reasonKey === 'STATUS_3L_EXTERNAL_BLANK_DETECTED');
  }
  if (sourceKey === 'SYSTEM') {
    // Legacy creation rows can establish WAITING/current non-DONE state, but
    // SYSTEM alone is never sufficient proof that a calibration was completed.
    if (reasonKey === 'CALIBRATION_CREATED' && statusKey !== 'DONE') return true;
    return statusKey !== 'DONE' && isResolvedStatus3LActor_(actor) && actorKey !== 'SYSTEM';
  }
  return isResolvedStatus3LActor_(actor) && actorKey !== 'SYSTEM';
}

function buildTrustedStatus3LStateIndex_(status3LLog) {
  const result = Object.create(null);
  if (!status3LLog || status3LLog.getLastRow() < 2 ||
      !historyHeadersReady_(status3LLog, UNIT_STATUS_3L_LOG_HEADERS)) return result;
  const map = getHistoryHeaderMap_(status3LLog, UNIT_STATUS_3L_LOG_HEADERS);
  const rows = status3LLog.getRange(2, 1, status3LLog.getLastRow() - 1, status3LLog.getLastColumn()).getValues();
  rows.forEach((row, offset) => {
    const key = calibrationCompletionKey_(
      row[map['NOMOR POLISI'] - 1], row[map['KALIBRASI KE'] - 1]
    );
    if (!key) return;
    const status3L = displayValue_(row[map['STATUS BARU'] - 1]);
    const source = displayValue_(row[map['SOURCE'] - 1]);
    const reason = map['REASON'] ? displayValue_(row[map['REASON'] - 1]) : '';
    const updatedBy = status3LActorDisplay_(displayValue_(row[map['USER EMAIL'] - 1]));
    if (!isTrustedStatus3LEvent_(updatedBy, source, reason, status3L)) return;
    const rawTimestamp = row[map['TIMESTAMP'] - 1];
    const parsed = rawTimestamp instanceof Date ? rawTimestamp : parseDateTimeInput_(rawTimestamp);
    const timestamp = parsed instanceof Date && !isNaN(parsed.getTime()) ? parsed : null;
    const event = {
      rowNumber: offset + 2,
      timestamp: timestamp,
      time: timestamp ? timestamp.getTime() : offset,
      updatedBy: updatedBy,
      previousStatus: displayValue_(row[map['STATUS SEBELUM'] - 1]),
      status3L: status3L,
      note: displayValue_(row[map['CATATAN'] - 1]),
      source: source,
      reason: reason,
      statusUnitBefore: displayValue_(row[map['STATUS UNIT SEBELUM'] - 1]),
      statusUnitAfter: displayValue_(row[map['STATUS UNIT SESUDAH'] - 1])
    };
    if (!result[key]) result[key] = { events: [], initial: null, done: null, latest: null };
    result[key].events.push(event);
  });
  Object.keys(result).forEach(key => {
    const bucket = result[key];
    bucket.events.sort((a, b) => a.time - b.time || a.rowNumber - b.rowNumber);
    bucket.initial = bucket.events.find(event =>
      normalizeLookupText_(event.reason) === 'CALIBRATION_CREATED' && cleanText_(event.status3L)
    ) || bucket.events.find(event => cleanText_(event.status3L)) || null;
    bucket.done = bucket.events.filter(event => normalizeLookupText_(event.status3L) === 'DONE').pop() || null;
    bucket.latest = bucket.events.length ? bucket.events[bucket.events.length - 1] : null;
  });
  return result;
}

function trustedStatus3LStateFor_(index, nopol, calibrationNumber) {
  const key = calibrationCompletionKey_(nopol, calibrationNumber);
  return key && index ? index[key] || null : null;
}

function isDirectTrustedFinalSnapshot_(row, map) {
  const type = normalizeLookupText_(getRowValue_(row, map, 'RECORD TYPE'));
  const lifecycle = normalizeLookupText_(getRowValue_(row, map, 'LIFECYCLE STATUS'));
  if (type !== 'FINAL_SNAPSHOT' || lifecycle !== 'LOCKED') return false;
  const source = normalizeLookupText_(getRowValue_(row, map, 'SOURCE'));
  const reason = normalizeLookupText_(getRowValue_(row, map, 'REASON'));
  const completedActor = calibrationHistoryActor_(getRowValue_(row, map, 'COMPLETED BY'));
  const status3LActor = calibrationHistoryActor_(getRowValue_(row, map, 'STATUS 3L UPDATED BY'));
  const actor = completedActor || status3LActor ||
    calibrationHistoryActor_(getRowValue_(row, map, 'USER EMAIL'));
  // v21.1 migration marker: a locked final whose two business attribution
  // fields were normalized to Team 3L is accepted as canonical completion
  // evidence. Raw technical provenance remains available in USER EMAIL/SOURCE.
  if (normalizeLookupText_(completedActor) === normalizeLookupText_(KM_STATUS_3L_SHEET_FALLBACK) &&
      normalizeLookupText_(status3LActor) === normalizeLookupText_(KM_STATUS_3L_SHEET_FALLBACK)) {
    return true;
  }
  if (isWeakStatus3LRepairReason_(reason)) return false;
  if (source === 'DASHBOARD' || source === 'SPREADSHEET' || source === 'SPREADSHEET_PRIMARY') return true;
  if (isTrustedStatus3LCorrectionReason_(reason)) {
    return isResolvedStatus3LActor_(actor) && normalizeLookupText_(actor) !== 'SYSTEM';
  }
  return reason === 'STATUS_3L_DONE' && isResolvedStatus3LActor_(actor) && normalizeLookupText_(actor) !== 'SYSTEM';
}

function buildTrustedFinalSnapshotIndex_(trustedStatusIndex) {
  const result = Object.create(null);
  const history = getUnitKmHistorySheet_(SpreadsheetApp.getActiveSpreadsheet());
  if (!history || history.getLastRow() < 2 || !historyHeadersReady_(history, UNIT_KM_LOG_HEADERS)) return result;
  const map = getHistoryHeaderMap_(history, UNIT_KM_LOG_HEADERS);
  const rows = history.getRange(2, 1, history.getLastRow() - 1, history.getLastColumn()).getValues();
  rows.forEach((row, offset) => {
    const type = normalizeLookupText_(getRowValue_(row, map, 'RECORD TYPE'));
    const lifecycle = normalizeLookupText_(getRowValue_(row, map, 'LIFECYCLE STATUS'));
    if (type !== 'FINAL_SNAPSHOT' || lifecycle !== 'LOCKED') return;
    const key = calibrationCompletionKey_(getRowValue_(row, map, 'NOMOR POLISI'), getRowValue_(row, map, 'KALIBRASI KE'));
    if (!key) return;
    const trustedDone = trustedStatusIndex && trustedStatusIndex[key] && trustedStatusIndex[key].done;
    if (!trustedDone && !isDirectTrustedFinalSnapshot_(row, map)) return;
    const rawTimestamp = getRowValue_(row, map, 'COMPLETED AT') || getRowValue_(row, map, 'TIMESTAMP');
    const timestamp = calibrationHistoryDate_(rawTimestamp);
    const candidate = {
      rowNumber: offset + 2,
      timestamp: timestamp,
      time: timestamp ? timestamp.getTime() : offset,
      updatedBy: calibrationHistoryActor_(getRowValue_(row, map, 'COMPLETED BY')) ||
        calibrationHistoryActor_(getRowValue_(row, map, 'USER EMAIL')),
      source: displayValue_(getRowValue_(row, map, 'SOURCE')),
      reason: displayValue_(getRowValue_(row, map, 'REASON'))
    };
    if (!result[key] || candidate.time >= result[key].time) result[key] = candidate;
  });
  return result;
}

function auditRepairSnapshotIntegrity_(auditContext) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const history = ss.getSheetByName(UNIT_KM_LOG_SHEET);
  const status3LLog = ss.getSheetByName(UNIT_STATUS_3L_LOG_SHEET);
  const trustedStatus = auditContext && auditContext.trustedStatus ||
    buildTrustedStatus3LStateIndex_(status3LLog);
  const trustedFinals = auditContext && auditContext.trustedFinals ||
    buildTrustedFinalSnapshotIndex_(trustedStatus);
  const report = {
    ready: true,
    lockedFinalSnapshots: 0,
    trustedFinalSnapshots: Object.keys(trustedFinals).length,
    invalidatedFinalSnapshots: 0,
    weakRepairFinalSnapshots: [],
    weakStatus3LRepairEvents: 0,
    legacyMasterStateReconciliationEvents: 0
  };
  if (history && history.getLastRow() > 1 && historyHeadersReady_(history, UNIT_KM_LOG_HEADERS)) {
    const map = getHistoryHeaderMap_(history, UNIT_KM_LOG_HEADERS);
    const rows = history.getRange(2, 1, history.getLastRow() - 1, history.getLastColumn()).getValues();
    rows.forEach((row, offset) => {
      const type = normalizeLookupText_(getRowValue_(row, map, 'RECORD TYPE'));
      const lifecycle = normalizeLookupText_(getRowValue_(row, map, 'LIFECYCLE STATUS'));
      if (lifecycle === 'INVALIDATED' || type === 'REPAIR_SNAPSHOT_INVALIDATED') report.invalidatedFinalSnapshots++;
      if (type !== 'FINAL_SNAPSHOT' || lifecycle !== 'LOCKED') return;
      report.lockedFinalSnapshots++;
      const nopol = normalizeDisplayNopol_(getRowValue_(row, map, 'NOMOR POLISI'));
      const sequence = parseCalibrationNumber_(getRowValue_(row, map, 'KALIBRASI KE'));
      const key = calibrationCompletionKey_(nopol, sequence);
      const source = displayValue_(getRowValue_(row, map, 'SOURCE'));
      const reason = displayValue_(getRowValue_(row, map, 'REASON'));
      const trustedDone = key && trustedStatus[key] && trustedStatus[key].done;
      if (trustedDone || isDirectTrustedFinalSnapshot_(row, map)) return;
      if (normalizeLookupText_(source) === 'SYSTEM_REPAIR' || isWeakStatus3LRepairReason_(reason)) {
        if (report.weakRepairFinalSnapshots.length < 200) report.weakRepairFinalSnapshots.push({
          row: offset + 2,
          nopol: nopol,
          calibrationNumber: sequence,
          source: source,
          reason: reason,
          completedAt: displayValue_(getRowValue_(row, map, 'COMPLETED AT')),
          completedBy: calibrationHistoryActor_(getRowValue_(row, map, 'COMPLETED BY'))
        });
      }
    });
  }
  if (status3LLog && status3LLog.getLastRow() > 1 && historyHeadersReady_(status3LLog, UNIT_STATUS_3L_LOG_HEADERS)) {
    const map = getHistoryHeaderMap_(status3LLog, UNIT_STATUS_3L_LOG_HEADERS);
    const rows = status3LLog.getRange(2, 1, status3LLog.getLastRow() - 1, status3LLog.getLastColumn()).getValues();
    rows.forEach(row => {
      const reason = displayValue_(row[map['REASON'] - 1]);
      if (isWeakStatus3LRepairReason_(reason)) report.weakStatus3LRepairEvents++;
      if (normalizeLookupText_(reason) === 'MASTER_STATE_RECONCILIATION') report.legacyMasterStateReconciliationEvents++;
    });
  }
  report.ready = report.weakRepairFinalSnapshots.length === 0;
  return report;
}

function status3LAuditEvidencePriority_(updatedBy, source, reason) {
  const actorDisplay = status3LActorDisplay_(updatedBy);
  const actorKey = normalizeLookupText_(actorDisplay);
  const sourceKey = normalizeLookupText_(source);
  const reasonKey = normalizeLookupText_(reason);
  if (!actorKey || isWeakStatus3LRepairReason_(reasonKey)) return 0;
  if (isTrustedStatus3LCorrectionReason_(reasonKey)) return actorKey === 'SYSTEM' ? 30 : 70;
  if (actorKey === normalizeLookupText_(KM_STATUS_3L_SHEET_FALLBACK)) {
    return (sourceKey === 'DASHBOARD' || sourceKey === 'SPREADSHEET' || sourceKey === 'SPREADSHEET_PRIMARY') ? 70 : 50;
  }
  if (/^SYSTEM(?:_|$)/.test(actorKey)) {
    // System/repair rows remain in the ledger but must not beat human evidence.
    if (sourceKey === 'AUTO_TRIGGER' && reasonKey === 'CALIBRATION_CREATED') return 20;
    return 10;
  }
  if (!isResolvedStatus3LActor_(updatedBy)) return 0;
  if (sourceKey === 'DASHBOARD' || sourceKey === 'SPREADSHEET' || sourceKey === 'SPREADSHEET_PRIMARY') return 100;
  if (sourceKey === 'EXTERNAL_SYNC') return 90;
  return 80;
}

function status3LEvidenceWins_(candidate, existing) {
  if (!existing) return true;
  // Event manual terbaru tetap otoritatif meskipun identitasnya disembunyikan.
  // Event SYSTEM repair tidak boleh menggantikan bukti manual.
  if (candidate.priority >= 60 && existing.priority >= 60 && candidate.time !== existing.time) {
    return candidate.time > existing.time;
  }
  return candidate.priority > existing.priority ||
    (candidate.priority === existing.priority && candidate.time >= existing.time);
}

/**
 * Ledger STATUS_3L_HISTORY adalah source of truth untuk siapa yang benar-benar
 * mengubah Status 3L. KM UPDATED BY / pengaju odometer tidak boleh digunakan
 * sebagai fallback attribution Status 3L.
 */
function buildStatus3LAuditEvidenceMap_(status3LLog) {
  const result = Object.create(null);
  if (!status3LLog || status3LLog.getLastRow() < 2 ||
      !historyHeadersReady_(status3LLog, UNIT_STATUS_3L_LOG_HEADERS)) return result;
  const map = getHistoryHeaderMap_(status3LLog, UNIT_STATUS_3L_LOG_HEADERS);
  const rows = status3LLog.getRange(2, 1, status3LLog.getLastRow() - 1, status3LLog.getLastColumn()).getValues();
  rows.forEach((row, offset) => {
    const key = status3LAuditKey_(
      row[map['NOMOR POLISI'] - 1],
      row[map['KALIBRASI KE'] - 1],
      row[map['STATUS BARU'] - 1]
    );
    if (!key) return;
    const rawTimestamp = row[map['TIMESTAMP'] - 1];
    const parsed = rawTimestamp instanceof Date ? rawTimestamp : parseDateTimeInput_(rawTimestamp);
    const timestamp = parsed instanceof Date && !isNaN(parsed.getTime()) ? parsed : '';
    const time = timestamp ? timestamp.getTime() : offset;
    const updatedBy = status3LActorDisplay_(displayValue_(row[map['USER EMAIL'] - 1]));
    const source = displayValue_(row[map['SOURCE'] - 1]);
    const reason = map['REASON'] ? displayValue_(row[map['REASON'] - 1]) : '';
    if (/NOTE_UPDATED/.test(normalizeLookupText_(reason))) return;
    const priority = status3LAuditEvidencePriority_(updatedBy, source, reason);
    if (priority <= 0) return;
    const existing = result[key];
    if (status3LEvidenceWins_({ priority: priority, time: time }, existing)) {
      result[key] = {
        time: time,
        priority: priority,
        timestamp: timestamp,
        updatedAt: timestamp ? timestamp.toISOString() : '',
        updatedBy: safeSheetText_(updatedBy),
        resolvedActor: isResolvedStatus3LActor_(updatedBy),
        source: source,
        reason: reason
      };
    }
  });
  return result;
}

function status3LAuditEvidenceFor_(auditMap, nopol, calibrationNumber, status3L) {
  if (!auditMap) return null;
  return auditMap[status3LAuditKey_(nopol, calibrationNumber, status3L)] || null;
}

function status3LActivityKey_(nopol, status3L) {
  const plate = normalizeNopol_(nopol);
  const statusKey = normalizeLookupText_(status3L);
  return plate && statusKey ? plate + '|' + statusKey : '';
}

/**
 * Secondary evidence untuk legacy bug v17/v18: pada jalur NOT_ELIGIBLE, edit
 * Status 3L sudah tercatat di KM_ACTIVITY_LOG tetapi STATUS_3L_HISTORY belum
 * sempat ditulis. Ini memungkinkan repair attribution tanpa menebak dari KM UPDATED BY.
 */
function buildStatus3LActivityEvidenceMap_(activitySheet) {
  const result = Object.create(null);
  if (!activitySheet || activitySheet.getLastRow() < 2 || activitySheet.getLastColumn() < 1) return result;
  // Evidence reader bersifat read-only; jangan menambah/menggeser header saat dashboard dibuka.
  const map = getHeaderMap_(activitySheet);
  const required = ['TIMESTAMP', 'NOMOR POLISI', 'USER EMAIL', 'ACTION', 'FIELD NAME', 'NEW VALUE'];
  if (required.some(header => !map[header])) return result;
  const validStatusKeys = new Set(
    ['WAITING', 'PROCESS', 'DONE', 'KENDALA'].concat(getStatus3LOptions_()).map(normalizeLookupText_)
  );
  const rows = activitySheet.getRange(2, 1, activitySheet.getLastRow() - 1, activitySheet.getLastColumn()).getValues();
  rows.forEach((row, offset) => {
    const action = normalizeLookupText_(getRowValue_(row, map, 'ACTION'));
    const fieldName = normalizeLookupText_(getRowValue_(row, map, 'FIELD NAME'));
    const status3L = displayValue_(getRowValue_(row, map, 'NEW VALUE'));
    const statusKey = normalizeLookupText_(status3L);
    const reliableAction = action === 'SHEET_STATUS_3L_NOT_ELIGIBLE' ||
      (action === 'SHEET_STATUS_3L_EDIT' && validStatusKeys.has(statusKey));
    if (!reliableAction || !/STATUS 3L/.test(fieldName) || !validStatusKeys.has(statusKey)) return;
    const updatedBy = status3LActorDisplay_(displayValue_(getRowValue_(row, map, 'USER EMAIL')));
    if (!isResolvedStatus3LActor_(updatedBy) || normalizeLookupText_(updatedBy) === 'SYSTEM') return;
    const key = status3LActivityKey_(getRowValue_(row, map, 'NOMOR POLISI'), status3L);
    if (!key) return;
    const details = safeParseHistoryJson_(getRowValue_(row, map, 'DETAILS JSON'));
    const rawTimestamp = details.editedAt || getRowValue_(row, map, 'TIMESTAMP');
    const parsed = rawTimestamp instanceof Date ? rawTimestamp : parseDateTimeInput_(rawTimestamp);
    const timestamp = parsed instanceof Date && !isNaN(parsed.getTime()) ? parsed : '';
    const time = timestamp ? timestamp.getTime() : offset;
    if (!result[key]) result[key] = [];
    result[key].push({
      time: time,
      timestamp: timestamp,
      updatedAt: timestamp ? timestamp.toISOString() : '',
      updatedBy: safeSheetText_(updatedBy),
      resolvedActor: true,
      source: 'KM_ACTIVITY_LOG',
      reason: action,
      calibrationNumber: parseCalibrationNumber_(details.calibrationNumber)
    });
  });
  Object.keys(result).forEach(key => result[key].sort((a, b) => a.time - b.time));
  return result;
}

function status3LActivityEvidenceFor_(activityMap, nopol, status3L, notBefore, calibrationNumber) {
  if (!activityMap) return null;
  const events = activityMap[status3LActivityKey_(nopol, status3L)] || [];
  if (!events.length) return null;
  const thresholdDate = notBefore instanceof Date ? notBefore : parseDateTimeInput_(notBefore);
  const threshold = thresholdDate instanceof Date && !isNaN(thresholdDate.getTime()) ? thresholdDate.getTime() : null;
  if (threshold === null) return null; // Tanpa batas siklus, bukti bisa berasal dari kalibrasi lama.
  for (let index = events.length - 1; index >= 0; index--) {
    const event = events[index];
    if (event.calibrationNumber && event.calibrationNumber !== parseCalibrationNumber_(calibrationNumber)) continue;
    if (event.timestamp && event.time >= threshold) return event;
  }
  return null;
}

function status3LAttributionEvidenceForRow_(auditMap, activityMap, row, indexes) {
  const direct = status3LAuditEvidenceFor_(
    auditMap, row[indexes.nopol], row[indexes.calibrationNumber], row[indexes.status3L]
  );
  const calibrationStart = indexes.kmUpdatedAt >= 0 ? row[indexes.kmUpdatedAt] : '';
  const fallback = status3LActivityEvidenceFor_(
    activityMap, row[indexes.nopol], row[indexes.status3L], calibrationStart, row[indexes.calibrationNumber]
  );
  let evidence = fallback && (!direct || fallback.time >= direct.time) ? fallback : (direct || null);

  // MASTER_UNIT tidak dipakai untuk MENEBak attribution yang hilang. Namun bila
  // primary state sudah menyimpan actor manusia yang valid dengan timestamp lebih
  // baru daripada ledger, repair tidak boleh menurunkannya ke event history lama.
  // Guard ini hanya mempertahankan attribution primary yang sudah ter-resolve.
  const primaryUpdatedBy = indexes.status3LUpdatedBy >= 0
    ? status3LActorDisplay_(displayValue_(row[indexes.status3LUpdatedBy])) : '';
  const primaryRawAt = indexes.status3LUpdatedAt >= 0 ? row[indexes.status3LUpdatedAt] : '';
  const primaryAt = primaryRawAt instanceof Date ? primaryRawAt : parseDateTimeInput_(primaryRawAt);
  const primaryValidAt = primaryAt instanceof Date && !isNaN(primaryAt.getTime()) ? primaryAt : null;
  if (isResolvedStatus3LActor_(primaryUpdatedBy) && normalizeLookupText_(primaryUpdatedBy) !== 'SYSTEM' && primaryValidAt) {
    const evidenceTime = evidence && evidence.timestamp instanceof Date ? evidence.timestamp.getTime() :
      (evidence && Number.isFinite(Number(evidence.time)) ? Number(evidence.time) : null);
    if (evidenceTime === null || primaryValidAt.getTime() > evidenceTime) {
      evidence = {
        time: primaryValidAt.getTime(),
        priority: 100,
        timestamp: primaryValidAt,
        updatedAt: primaryValidAt.toISOString(),
        updatedBy: safeSheetText_(primaryUpdatedBy),
        resolvedActor: true,
        source: 'MASTER_PRIMARY_ATTRIBUTION',
        reason: 'PRIMARY_ATTRIBUTION_NEWER_THAN_LEDGER'
      };
    }
  }
  return evidence;
}

/**
 * Sinkronisasi snapshot audit MASTER_UNIT dari STATUS_3L_HISTORY.
 * Berbeda dari implementasi lama, evidence history yang cocok boleh mengganti
 * nilai lama yang salah (mis. masih berisi KM UPDATED BY/pengaju).
 */
function backfillStatus3LTimestamps_(masterSheet, indexes, status3LLog) {
  const report = { changed: 0, matched: 0, unresolved: 0, skippedConcurrentChanges: 0, mismatches: [] };
  if (!masterSheet || masterSheet.getLastRow() < 2 || indexes.status3LUpdatedAt < 0 ||
      indexes.status3LUpdatedBy !== indexes.status3LUpdatedAt + 1 || !status3LLog) return report;
  const auditMap = buildStatus3LAuditEvidenceMap_(status3LLog);
  const activityMap = buildStatus3LActivityEvidenceMap_(
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.LOG)
  );
  const rowCount = masterSheet.getLastRow() - 1;
  const rows = masterSheet.getRange(2, 1, rowCount, masterSheet.getLastColumn()).getValues();
  const auditValues = rows.map((row, offset) => {
    const existingTimestamp = row[indexes.status3LUpdatedAt];
    const currentStatus = displayValue_(row[indexes.status3L]);
    const existingUpdatedBy = normalizeDoneStatus3LActor_(currentStatus,
      displayValue_(row[indexes.status3LUpdatedBy]));
    const evidence = status3LAttributionEvidenceForRow_(auditMap, activityMap, row, indexes);
    if (!evidence) {
      if (cleanText_(currentStatus)) report.unresolved++;
      return [existingTimestamp, existingUpdatedBy];
    }
    report.matched++;
    // Bila event historinya sendiri tidak memiliki editor yang dapat diresolusikan,
    // jangan menebak dari KM UPDATED BY atau akun pengaju.
    if (!evidence.resolvedActor) {
      report.unresolved++;
      // Repair tidak boleh mengalahkan editor pada primary state dengan SYSTEM/marker.
      return [existingTimestamp, existingUpdatedBy];
    }
    const existingDate = parseDateTimeInput_(existingTimestamp);
    if (isResolvedStatus3LActor_(existingUpdatedBy) && existingDate &&
        (!evidence.timestamp || existingDate.getTime() > evidence.timestamp.getTime())) {
      return [existingTimestamp, existingUpdatedBy];
    }
    const nextTimestamp = evidence.timestamp || existingTimestamp || '';
    const nextUpdatedBy = normalizeDoneStatus3LActor_(currentStatus, evidence.updatedBy);
    const currentTime = existingTimestamp instanceof Date ? existingTimestamp.getTime() :
      ((parseDateTimeInput_(existingTimestamp) || {}).getTime ? parseDateTimeInput_(existingTimestamp).getTime() : 0);
    const nextTime = nextTimestamp instanceof Date ? nextTimestamp.getTime() :
      ((parseDateTimeInput_(nextTimestamp) || {}).getTime ? parseDateTimeInput_(nextTimestamp).getTime() : 0);
    const mismatch = normalizeLookupText_(existingUpdatedBy) !== normalizeLookupText_(nextUpdatedBy) || currentTime !== nextTime;
    if (mismatch) {
      if (report.mismatches.length < 100) report.mismatches.push({
        row: offset + 2,
        nopol: normalizeDisplayNopol_(row[indexes.nopol]),
        status3L: currentStatus,
        calibrationNumber: parseCalibrationNumber_(row[indexes.calibrationNumber]),
        previousUpdatedBy: existingUpdatedBy,
        correctedUpdatedBy: nextUpdatedBy
      });
    }
    return [nextTimestamp, safeSheetText_(nextUpdatedBy)];
  });
  auditValues.forEach((values, offset) => {
    const old = rows[offset];
    if (String(old[indexes.status3LUpdatedAt]) === String(values[0]) &&
        String(old[indexes.status3LUpdatedBy]) === String(values[1])) return;
    // Native edits tidak ikut ScriptLock: jangan timpa row yang berubah selama audit.
    const live = masterSheet.getRange(offset + 2, 1, 1, masterSheet.getLastColumn()).getValues()[0];
    if (['nopol', 'calibrationNumber', 'status3L', 'status3LUpdatedAt', 'status3LUpdatedBy']
        .some(field => String(live[indexes[field]]) !== String(old[indexes[field]]))) {
      report.skippedConcurrentChanges++;
      return;
    }
    masterSheet.getRange(offset + 2, indexes.status3LUpdatedAt + 1, 1, 2).setValues([values]);
    report.changed++;
  });
  return report;
}


function calibrationCompletionKey_(nopol, calibrationNumber) {
  const plate = normalizeNopol_(nopol);
  const sequence = parseCalibrationNumber_(calibrationNumber);
  return plate && sequence ? plate + '|' + sequence : '';
}

function buildCompletedCalibrationKeySet_() {
  const result = new Set();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trustedStatus = buildTrustedStatus3LStateIndex_(ss.getSheetByName(UNIT_STATUS_3L_LOG_SHEET));
  Object.keys(trustedStatus).forEach(key => {
    if (trustedStatus[key] && trustedStatus[key].done) result.add(key);
  });
  const finals = buildTrustedFinalSnapshotIndex_(trustedStatus);
  Object.keys(finals).forEach(key => result.add(key));
  return result;
}

function isCalibrationCompleted_(nopol, calibrationNumber, completedKeys) {
  const key = calibrationCompletionKey_(nopol, calibrationNumber);
  if (!key) return false;
  const keys = completedKeys instanceof Set ? completedKeys : buildCompletedCalibrationKeySet_();
  return keys.has(key);
}

// Maintenance is not a business transition. Native edits may already be visible
// while their onEdit history/timestamp is still pending. Never replay history
// over a populated current status, even when the history timestamp is newer.
function canReplayStatus3LHistory_(currentStatus, currentTimestamp, latest) {
  if (cleanText_(currentStatus) || !latest || !latest.timestamp) return false;
  const currentAt = parseDateTimeInput_(currentTimestamp);
  // An explicit blank cannot safely be distinguished from a pending native edit.
  // Restoring historical completion belongs to the explicit recovery operation.
  return !currentAt && normalizeLookupText_(latest.status3L) === 'DONE';
}

function masterWorkflowRowUnchanged_(before, live, indexes) {
  return ['nopol', 'calibrationNumber', 'statusUnit', 'status3L',
    'status3LUpdatedAt', 'status3LUpdatedBy', 'status3LNote', 'kmUpdatedAt']
    .every(field => indexes[field] < 0 || indexes[field] === undefined ||
      String(before[indexes[field]] || '') === String(live[indexes[field]] || ''));
}

function reconcileMasterUnitWorkflowState_(sheet, indexes) {
  const report = {
    changed: 0,
    status3LChanged: 0,
    statusUnitChanged: 0,
    attributionAligned: 0,
    unsupportedDone: 0,
    unresolved: 0,
    skippedConcurrentChanges: 0,
    rows: []
  };
  if (!sheet || sheet.getLastRow() < 2) return report;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trustedIndex = buildTrustedStatus3LStateIndex_(ss.getSheetByName(UNIT_STATUS_3L_LOG_SHEET));
  const completedKeys = buildCompletedCalibrationKeySet_();
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const now = new Date();
  const status3LEntries = [];
  const statusUnitEntries = [];

  rows.forEach((row, offset) => {
    const nopol = normalizeDisplayNopol_(row[indexes.nopol]);
    const sequence = parseCalibrationNumber_(row[indexes.calibrationNumber]);
    if (!nopol || (!sequence && normalizeLookupText_(row[indexes.status3L]) !== 'DONE')) return;
    const key = calibrationCompletionKey_(nopol, sequence);
    const bucket = trustedIndex[key] || null;
    const latest = bucket && bucket.latest || null;
    const currentStatus3L = displayValue_(row[indexes.status3L]);
    const currentUnit = displayValue_(row[indexes.statusUnit]);
    const currentStatusKey = normalizeLookupText_(currentStatus3L);
    const currentUnitKey = normalizeLookupText_(currentUnit);
    let desiredStatus3L = currentStatus3L;
    let desiredUnit = currentUnit;
    let desiredUpdatedAt = row[indexes.status3LUpdatedAt];
    let desiredUpdatedBy = displayValue_(row[indexes.status3LUpdatedBy]);
    let desiredNote = displayValue_(row[indexes.status3LNote]);
    let reason = '';

    if (currentStatusKey === 'DONE') {
      // Enforce the current business state even if the asynchronous ledger failed.
      // Missing independent evidence remains an audit issue, never a rollback.
      if (!completedKeys.has(key)) report.unsupportedDone++;
      desiredStatus3L = 'DONE';
      desiredUnit = 'Ready to Depart';
      desiredUpdatedBy = preferredDoneStatus3LActor_(
        desiredUpdatedBy, bucket && bucket.done ? bucket.done.updatedBy : ''
      );
      reason = 'STATUS_3L_DONE_ATTRIBUTION_ALIGNED';
    } else if (canReplayStatus3LHistory_(currentStatus3L, desiredUpdatedAt, latest) &&
        currentUnitKey !== normalizeLookupText_('On Delivery')) {
      desiredStatus3L = latest.status3L;
      if (latest.timestamp) desiredUpdatedAt = latest.timestamp;
      if (latest.updatedBy) desiredUpdatedBy = latest.updatedBy;
      if (latest.note && normalizeLookupText_(latest.status3L) === normalizeLookupText_(currentStatus3L)) {
        desiredNote = latest.note;
      }
      const eventUnit = latest.statusUnitAfter;
      if (UNIT_OPERATIONAL_STATUSES.some(status => normalizeLookupText_(status) === normalizeLookupText_(eventUnit))) {
        desiredUnit = eventUnit;
      }
      if (normalizeLookupText_(desiredStatus3L) === 'DONE') {
        desiredUnit = 'Ready to Depart';
      }
      reason = 'STATUS_3L_MASTER_STATE_RECONCILED_FROM_TRUSTED_HISTORY';
    } else {
      return;
    }

    if (normalizeLookupText_(desiredStatus3L) === 'DONE') {
      desiredUnit = 'Ready to Depart';
      desiredUpdatedBy = preferredDoneStatus3LActor_(
        desiredUpdatedBy, bucket && bucket.done ? bucket.done.updatedBy : ''
      );
    }

    const statusChanged = normalizeLookupText_(desiredStatus3L) !== currentStatusKey;
    const unitChanged = normalizeLookupText_(desiredUnit) !== currentUnitKey;
    const atChanged = String(desiredUpdatedAt || '') !== String(row[indexes.status3LUpdatedAt] || '');
    const byChanged = desiredUpdatedBy !== displayValue_(row[indexes.status3LUpdatedBy]);
    if (!statusChanged && !unitChanged && !atChanged && !byChanged) return;

    const live = sheet.getRange(offset + 2, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!masterWorkflowRowUnchanged_(row, live, indexes)) {
      report.skippedConcurrentChanges++;
      return;
    }

    const updateId = Utilities.getUuid();
    if (statusChanged || atChanged || byChanged) {
      sheet.getRange(offset + 2, indexes.status3L + 1, 1, 4).setValues([[
        safeSheetText_(desiredStatus3L), desiredUpdatedAt || now,
        safeSheetText_(desiredUpdatedBy || 'SYSTEM'), safeSheetText_(desiredNote)
      ]]);
      report.status3LChanged += statusChanged ? 1 : 0;
      report.attributionAligned += (!statusChanged && (atChanged || byChanged)) ? 1 : 0;
      if (statusChanged) status3LEntries.push({
        updateId: updateId,
        timestamp: now,
        email: normalizeDoneStatus3LActor_(desiredStatus3L, desiredUpdatedBy || 'SYSTEM'),
        nopol: nopol,
        calibrationNumber: sequence,
        previousStatus: currentStatus3L,
        status3L: desiredStatus3L,
        note: 'Rekonsiliasi state MASTER_UNIT dari evidence history tepercaya; MASTER_UNIT tidak digunakan sebagai sumber bukti.',
        source: 'SYSTEM_REPAIR',
        reason: reason,
        previousStatusUnit: currentUnit,
        statusUnit: desiredUnit
      });
      else if (byChanged && normalizeLookupText_(desiredStatus3L) === 'DONE') status3LEntries.push({
        updateId: updateId,
        timestamp: now,
        email: KM_STATUS_3L_SHEET_FALLBACK,
        nopol: nopol,
        calibrationNumber: sequence,
        previousStatus: currentStatus3L,
        status3L: desiredStatus3L,
        note: 'Normalisasi attribution Status 3L DONE ke Team 3L tanpa mengubah status bisnis.',
        source: 'ADMIN_NORMALIZATION',
        reason: 'STATUS_3L_DONE_ATTRIBUTION_NORMALIZED_TO_TEAM_3L',
        previousStatusUnit: currentUnit,
        statusUnit: desiredUnit
      });
    }
    if (unitChanged) {
      sheet.getRange(offset + 2, indexes.statusUnit + 1).setValue(safeSheetText_(desiredUnit));
      report.statusUnitChanged++;
      statusUnitEntries.push({
        updateId: updateId,
        timestamp: now,
        email: 'SYSTEM',
        nopol: nopol,
        calibrationNumber: sequence,
        previousStatus: currentUnit,
        statusUnit: desiredUnit,
        note: 'Rekonsiliasi Status Unit dari workflow/evidence Status 3L tepercaya.',
        source: 'SYSTEM_REPAIR',
        reason: reason
      });
    }
    report.changed++;
    if (report.rows.length < 100) report.rows.push({
      row: offset + 2, nopol: nopol, calibrationNumber: sequence,
      previousStatus3L: currentStatus3L, status3L: desiredStatus3L,
      previousStatusUnit: currentUnit, statusUnit: desiredUnit, reason: reason
    });
  });

  appendStatus3LHistoryRows_(status3LEntries);
  appendStatusUnitHistoryRows_(statusUnitEntries);
  SpreadsheetApp.flush();
  return report;
}



function auditCalibrationSummaryIntegrity_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  const report = {
    totalUnits: 0,
    status3LMismatchCount: 0,
    statusUnitMismatchCount: 0,
    statusUnitHistoryMissingCount: 0,
    doneMissingFinalSnapshotCount: 0,
    unsupportedDoneCount: 0,
    samples: []
  };
  if (!master || master.getLastRow() < 2) return report;
  const headers = master.getRange(1, 1, 1, master.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  if (indexes.nopol < 0 || indexes.statusUnit < 0 || indexes.status3L < 0 || indexes.calibrationNumber < 0) return report;

  // v21: dashboard/administration audit must compare MASTER_UNIT against
  // independent trusted evidence, never against raw SYSTEM_REPAIR mirrors.
  const trustedStatus = buildTrustedStatus3LStateIndex_(ss.getSheetByName(UNIT_STATUS_3L_LOG_SHEET));
  const trustedFinals = buildTrustedFinalSnapshotIndex_(trustedStatus);
  const latestUnit = buildLatestStatusUnitMap_();
  const rows = master.getRange(2, 1, master.getLastRow() - 1, master.getLastColumn()).getValues();
  rows.forEach((row, offset) => {
    const nopol = normalizeDisplayNopol_(row[indexes.nopol]);
    if (!nopol) return;
    report.totalUnits++;
    const sequence = parseCalibrationNumber_(row[indexes.calibrationNumber]);
    const status3L = displayValue_(row[indexes.status3L]);
    const statusUnit = displayValue_(row[indexes.statusUnit]);
    const calibrationKey = calibrationCompletionKey_(nopol, sequence);
    const current3LKey = normalizeLookupText_(status3L);
    const bucket = calibrationKey ? trustedStatus[calibrationKey] || null : null;
    const trustedLatest = bucket && bucket.latest || null;
    const history3L = trustedLatest ? trustedLatest.status3L : undefined;
    const unitKey = normalizeNopol_(nopol);
    const historyUnit = Object.prototype.hasOwnProperty.call(latestUnit, unitKey) ? latestUnit[unitKey] : undefined;
    const representBlank3L = !current3LKey && normalizeLookupText_(statusUnit) === normalizeLookupText_('On Delivery');
    const hasTrustedEvidence = Boolean(trustedLatest);
    const status3LMismatch = Boolean(sequence && hasTrustedEvidence &&
      normalizeLookupText_(history3L) !== current3LKey &&
      !(representBlank3L && !normalizeLookupText_(history3L)));
    const unsupportedDone = Boolean(sequence && current3LKey === 'DONE' && !(bucket && bucket.done) && !trustedFinals[calibrationKey]);
    const statusUnitHistoryMissing = Boolean(cleanText_(statusUnit) && historyUnit === undefined);
    const statusUnitMismatch = Boolean(cleanText_(statusUnit) && historyUnit !== undefined &&
      normalizeLookupText_(historyUnit) !== normalizeLookupText_(statusUnit));
    const missingFinal = Boolean(calibrationKey && current3LKey === 'DONE' && (bucket && bucket.done) && !trustedFinals[calibrationKey]);
    if (status3LMismatch) report.status3LMismatchCount++;
    if (statusUnitMismatch) report.statusUnitMismatchCount++;
    if (statusUnitHistoryMissing) report.statusUnitHistoryMissingCount++;
    if (missingFinal) report.doneMissingFinalSnapshotCount++;
    if (unsupportedDone) report.unsupportedDoneCount++;
    if ((status3LMismatch || statusUnitMismatch || missingFinal || unsupportedDone) && report.samples.length < 100) {
      report.samples.push({
        row: offset + 2, nopol: nopol, calibrationNumber: sequence,
        masterStatus3L: status3L, trustedStatus3L: history3L === undefined ? '' : history3L,
        masterStatusUnit: statusUnit, historyStatusUnit: historyUnit === undefined ? '' : historyUnit,
        statusUnitHistoryMissing: statusUnitHistoryMissing,
        missingTrustedFinalSnapshot: missingFinal,
        unsupportedDone: unsupportedDone
      });
    }
  });
  report.ready = report.status3LMismatchCount === 0 && report.statusUnitMismatchCount === 0 &&
    report.doneMissingFinalSnapshotCount === 0 && report.unsupportedDoneCount === 0;
  return report;
}

function auditCalibrationSummaryIntegrityFromSheet() {
  requireSheetMaintenanceAdmin_();
  return { success: true, audit: auditCalibrationSummaryIntegrity_() };
}

function repairCalibrationSummaryIntegrityFromSheet() {
  const actor = requireSheetMaintenanceAdmin_();
  KM_REQUEST_ACTOR_ = actor;
  const before = auditCalibrationSummaryIntegrity_();
  const repair = setupMasterUnitKmDatabase_();
  const after = auditCalibrationSummaryIntegrity_();
  return { success: true, before: before, repair: repair, after: after };
}

function auditStatus3LUpdatedByIntegrity() {
  const actor = requirePermission_('ADMIN_SETUP', '', { source: 'SHEET' });
  KM_REQUEST_ACTOR_ = actor;
  const audit = auditMasterUnitKmDatabase_();
  return {
    success: true,
    mismatchCount: (audit.status3LAttributionMismatchRows || []).length,
    unresolvedCount: (audit.status3LAttributionUnresolvedRows || []).length,
    mismatches: audit.status3LAttributionMismatchRows || [],
    unresolved: audit.status3LAttributionUnresolvedRows || []
  };
}

function repairStatus3LUpdatedByFromHistory() {
  const actor = requirePermission_('ADMIN_SETUP', '', { source: 'SHEET' });
  KM_REQUEST_ACTOR_ = actor;
  const lock = acquireScriptLock_(30000, 'repair attribution Status 3L');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const master = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
    const history = ss.getSheetByName(UNIT_STATUS_3L_LOG_SHEET);
    if (!master || master.getLastRow() < 2) throw new Error('MASTER_UNIT belum tersedia.');
    if (!history || !historyHeadersReady_(history, UNIT_STATUS_3L_LOG_HEADERS)) {
      throw new Error('STATUS_3L_HISTORY belum siap; attribution tidak dapat direkonstruksi dengan aman.');
    }
    const headers = master.getRange(1, 1, 1, master.getLastColumn()).getDisplayValues()[0];
    const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
    if (indexes.nopol < 0 || indexes.calibrationNumber < 0 || indexes.status3L < 0 ||
        indexes.status3LUpdatedAt < 0 || indexes.status3LUpdatedBy < 0) {
      throw new Error('Kolom workflow Status 3L belum lengkap.');
    }
    const report = backfillStatus3LTimestamps_(master, indexes, history);
    SpreadsheetApp.flush();
    invalidateAppCaches_();
    logSystemActivity_('', '', 'STATUS_3L_ATTRIBUTION_REPAIRED', 'STATUS 3L UPDATED BY', '', '',
      'Rekonsiliasi attribution dari STATUS_3L_HISTORY: ' + JSON.stringify({
        changed: report.changed, matched: report.matched, unresolved: report.unresolved
      }), actor, true);
    return { success: true, report: report, audit: auditStatus3LUpdatedByIntegrity() };
  } finally {
    lock.releaseLock();
  }
}

function auditStatus3LDoneRegression() {
  requirePermission_('ADMIN_SETUP', '', { source: 'SHEET' });
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNIT_STATUS_3L_LOG_SHEET);
  if (!sheet || sheet.getLastRow() < 2 || !historyHeadersReady_(sheet, UNIT_STATUS_3L_LOG_HEADERS)) {
    return { success: true, count: 0, rows: [] };
  }
  const map = getHistoryHeaderMap_(sheet, UNIT_STATUS_3L_LOG_HEADERS);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const regressions = [];
  rows.forEach((row, offset) => {
    const previous = displayValue_(row[map['STATUS SEBELUM'] - 1]);
    const next = displayValue_(row[map['STATUS BARU'] - 1]);
    if (normalizeLookupText_(previous) !== 'DONE' || normalizeLookupText_(next) === 'DONE') return;
    regressions.push({
      row: offset + 2,
      timestamp: dateToIso_(row[map['TIMESTAMP'] - 1]),
      nopol: displayValue_(row[map['NOMOR POLISI'] - 1]),
      calibrationNumber: parseCalibrationNumber_(row[map['KALIBRASI KE'] - 1]),
      previousStatus3L: previous,
      status3L: next,
      updatedBy: status3LActorDisplay_(displayValue_(row[map['USER EMAIL'] - 1])),
      source: displayValue_(row[map['SOURCE'] - 1]),
      reason: map['REASON'] ? displayValue_(row[map['REASON'] - 1]) : '',
      previousStatusUnit: map['STATUS UNIT SEBELUM'] ? displayValue_(row[map['STATUS UNIT SEBELUM'] - 1]) : '',
      statusUnit: map['STATUS UNIT SESUDAH'] ? displayValue_(row[map['STATUS UNIT SESUDAH'] - 1]) : ''
    });
  });
  return { success: true, count: regressions.length, rows: regressions.slice(-200) };
}

function repairCompletedStatus3LFromHistory() {
  return repairCompletedStatus3LFromHistory_();
}

function repairCompletedStatus3LFromHistory_() {
  const actor = requirePermission_('ADMIN_SETUP', '', { source: 'SHEET' });
  KM_REQUEST_ACTOR_ = actor;
  const lock = acquireScriptLock_(60000, 'repair Status 3L DONE dari trusted history');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const master = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
    const status3LSheet = ss.getSheetByName(UNIT_STATUS_3L_LOG_SHEET);
    if (!master || master.getLastRow() < 2 || !status3LSheet || status3LSheet.getLastRow() < 2) {
      return { success: true, repaired: 0, message: 'Tidak ada data/evidence tepercaya yang perlu diperbaiki.' };
    }
    const headers = master.getRange(1, 1, 1, master.getLastColumn()).getDisplayValues()[0];
    const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
    const trustedStatus = buildTrustedStatus3LStateIndex_(status3LSheet);
    const trustedFinals = buildTrustedFinalSnapshotIndex_(trustedStatus);
    const rows = master.getRange(2, 1, master.getLastRow() - 1, master.getLastColumn()).getValues();
    const repaired = [];
    const statusEntries = [];
    const unitEntries = [];
    rows.forEach((row, offset) => {
      const nopol = normalizeDisplayNopol_(row[indexes.nopol]);
      const sequence = parseCalibrationNumber_(row[indexes.calibrationNumber]);
      const key = calibrationCompletionKey_(nopol, sequence);
      if (!key) return;
      const bucket = trustedStatus[key] || null;
      const done = bucket && bucket.done || null;
      const finalSnapshot = trustedFinals[key] || null;
      // DONE may only be restored when there is independent trusted DONE evidence.
      // A FINAL_SNAPSHOT alone is insufficient if it came from legacy SYSTEM_REPAIR.
      if (!done && !finalSnapshot) return;
      const evidence = done || finalSnapshot;
      const current = displayValue_(row[indexes.status3L]);
      const currentUnit = displayValue_(row[indexes.statusUnit]);
      const statusNeedsRepair = normalizeLookupText_(current) !== 'DONE';
      const unitNeedsRepair = normalizeLookupText_(currentUnit) !== normalizeLookupText_('Ready to Depart');
      const currentActor = indexes.status3LUpdatedBy >= 0
        ? status3LActorDisplay_(displayValue_(row[indexes.status3LUpdatedBy])) : '';
      const repairedActor = preferredDoneStatus3LActor_(currentActor, evidence.updatedBy || '');
      const attributionNeedsRepair = indexes.status3LUpdatedBy >= 0 &&
        normalizeLookupText_(currentActor) !== normalizeLookupText_(repairedActor);
      if (!statusNeedsRepair && !unitNeedsRepair && !attributionNeedsRepair) return;
      const rowNumber = offset + 2;
      const timestamp = evidence.timestamp || new Date();
      const updatedBy = repairedActor;
      const live = master.getRange(rowNumber, 1, 1, master.getLastColumn()).getValues()[0];
      if (!masterWorkflowRowUnchanged_(row, live, indexes)) return;
      if (statusNeedsRepair) {
        master.getRange(rowNumber, indexes.status3L + 1, 1, 4)
          .setValues([['DONE', timestamp, safeSheetText_(updatedBy), displayValue_(row[indexes.status3LNote])]]);
        statusEntries.push({
          updateId: Utilities.getUuid(), timestamp: new Date(), email: 'SYSTEM', nopol: nopol,
          calibrationNumber: sequence, previousStatus: current, status3L: 'DONE',
          note: 'Repair evidence-first: Status 3L dipulihkan dari trusted DONE/final evidence.',
          source: 'SYSTEM_REPAIR', reason: 'STATUS_3L_MASTER_STATE_RECONCILED_FROM_TRUSTED_HISTORY',
          previousStatusUnit: currentUnit, statusUnit: 'Ready to Depart'
        });
      }
      if (unitNeedsRepair) {
        master.getRange(rowNumber, indexes.statusUnit + 1).setValue('Ready to Depart');
        unitEntries.push({
          updateId: Utilities.getUuid(), timestamp: new Date(), email: 'SYSTEM', nopol: nopol,
          calibrationNumber: sequence, previousStatus: currentUnit, statusUnit: 'Ready to Depart',
          note: 'Ready to Depart dipulihkan hanya karena terdapat trusted DONE evidence.',
          source: 'SYSTEM_REPAIR', reason: 'STATUS_3L_MASTER_STATE_RECONCILED_FROM_TRUSTED_HISTORY'
        });
      }
      if (indexes.status3LUpdatedBy >= 0 && attributionNeedsRepair) {
        master.getRange(rowNumber, indexes.status3LUpdatedBy + 1).setValue(safeSheetText_(updatedBy));
      }
      repaired.push({ row: rowNumber, nopol: nopol, calibrationNumber: sequence,
        previousStatus3L: current, previousStatusUnit: currentUnit });
    });
    appendStatus3LHistoryRows_(statusEntries);
    appendStatusUnitHistoryRows_(unitEntries);
    SpreadsheetApp.flush();
    invalidateAppCaches_();
    return { success: true, repaired: repaired.length, rows: repaired };
  } finally {
    lock.releaseLock();
  }
}

function formatAllSheetDateColumns_(ss) {
  ss.getSheets().forEach(sheet => {
    if (!sheet.getLastColumn() || sheet.getMaxRows() < 2) return;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const dataRows = sheet.getMaxRows() - 1;
    headers.forEach((header, index) => {
      const key = cleanText_(header).toUpperCase().replace(/\s+/g, ' ');
      if (!key) return;
      const dateOnly = /(^| )(TANGGAL|TGL|DATE)( |$)/.test(key);
      const dateTime = key === 'TIMESTAMP' || / AT$/.test(key) || / UNTIL$/.test(key);
      if (dateOnly || dateTime) {
        trySetNumberFormat_(
          sheet.getRange(2, index + 1, dataRows, 1),
          dateTime ? 'dd mmmm yyyy, HH.mm.ss' : 'dd mmmm yyyy'
        );
      }
    });
  });
}



/** Shared cache, sheet, normalization, and date utilities. */
function getConfigParameter_(key, defaultValue) {
  const parameters = getConfigParameters_();
  const value = Object.prototype.hasOwnProperty.call(parameters, key)
    ? parameters[key]
    : '';
  return value || defaultValue;
}

function getConfigParameters_() {
  if (KM_REQUEST_CONFIG_PARAMETERS_) return Object.assign({}, KM_REQUEST_CONFIG_PARAMETERS_);
  const cached = getCachedJson_('config-parameters');
  if (cached) {
    KM_REQUEST_CONFIG_PARAMETERS_ = cached;
    return Object.assign({}, cached);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(KM_APP.SHEETS.CONFIG);
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 2) return {};

  const result = getExistingParameters_(sheet);
  putCachedJson_('config-parameters', result, 120);
  KM_REQUEST_CONFIG_PARAMETERS_ = result;
  return Object.assign({}, result);
}

function logSystemActivity_(
  ticket,
  nopol,
  action,
  fieldName,
  oldValue,
  newValue,
  note,
  actor,
  lockAlreadyHeld
) {
  logSystemActivityBatch_([
    {
      ticket: ticket,
      nopol: nopol,
      action: action,
      fieldName: fieldName,
      oldValue: oldValue,
      newValue: newValue,
      note: note,
      actor: actor
    }
  ], actor, lockAlreadyHeld);
}

function logSystemActivityBatch_(entries, actorOverride, lockAlreadyHeld) {
  if (!entries || !entries.length) return;
  if (lockAlreadyHeld) return writeSystemActivityBatch_(entries, actorOverride);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return writeSystemActivityBatch_(entries, actorOverride);
  } finally {
    lock.releaseLock();
  }
}

function writeSystemActivityBatch_(entries, actorOverride) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, KM_APP.SHEETS.LOG);

  ensureHeaders_(sheet, KM_APP.LOG_HEADERS);

  const now = new Date();
  const fallbackActor = actorOverride || getCurrentUserContext_({ allowUnauthorized: true });
  const stamp = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    'yyyyMMddHHmmssSSS'
  );

  const map = getHeaderMap_(sheet);
  const rows = entries.map((entry, index) => {
    const actor = entry.actor || fallbackActor || {};
    const rawSource = normalizeLookupText_(entry.source || actor.source || 'SYSTEM');
    const source = rawSource === 'WEB_APP' ? 'DASHBOARD'
      : rawSource === 'SHEET' ? 'SPREADSHEET'
      : (rawSource || 'SYSTEM');
    const row = new Array(sheet.getLastColumn()).fill('');
    setRowValue_(row, map, 'LOG ID', 'LOG-' + stamp + '-' + String(index + 1).padStart(4, '0'));
    setRowValue_(row, map, 'TIMESTAMP', now);
    setRowValue_(row, map, 'NOMOR POLISI', safeSheetText_(entry.nopol));
    setRowValue_(row, map, 'USER EMAIL', safeSheetText_(actor.email || getCurrentUserEmail_()));
    setRowValue_(row, map, 'ACTION', safeSheetText_(entry.action));
    setRowValue_(row, map, 'FIELD NAME', safeSheetText_(entry.fieldName));
    setRowValue_(row, map, 'OLD VALUE', safeSheetText_(entry.oldValue));
    setRowValue_(row, map, 'NEW VALUE', safeSheetText_(entry.newValue));
    setRowValue_(row, map, 'NOTE', safeSheetText_(entry.note));
    setRowValue_(row, map, 'USER ROLE', safeSheetText_((actor.roles || []).join(', ')));
    setRowValue_(row, map, 'SOURCE', safeSheetText_(source));
    setRowValue_(row, map, 'DETAILS JSON', safeSheetText_(entry.details ? JSON.stringify(entry.details) : ''));
    return row;
  });

  ensureSheetCapacity_(sheet, rows.length);

  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      rows.length,
      sheet.getLastColumn()
    )
    .setValues(rows);
}

function formatConfigSheet_(sheet) {
  sheet.setFrozenRows(1);
  sheet.getRange('A1:D1')
    .setBackground('#1F6B45')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange('E1:F1')
    .setBackground('#334155')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 170);
  sheet.setColumnWidth(5, 170);
  sheet.setColumnWidth(6, 220);
  sheet.setRowHeight(1, 38);
}

function formatLogSheet_(sheet) {
  const lastCol = sheet.getLastColumn();

  sheet
    .getRange(1, 1, 1, lastCol)
    .setBackground('#374151')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  sheet.setFrozenRows(1);

  const map = getHeaderMap_(sheet);

  if (map['TIMESTAMP']) {
    trySetNumberFormat_(
      sheet.getRange(
        2,
        map['TIMESTAMP'],
        Math.max(sheet.getMaxRows() - 1, 1),
        1
      ),
      'dd mmmm yyyy, HH.mm.ss'
    );
  }

  if (!sheet.getFilter()) {
    sheet
      .getRange(
        1,
        1,
        sheet.getMaxRows(),
        lastCol
      )
      .createFilter();
  }
}

function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  return sheet;
}

function getCachedJson_(key) {
  if (typeof CacheService === 'undefined') return null;
  try {
    const value = CacheService.getScriptCache().get('km-app-v8:' + key);
    if (!value) return null;
    const json = value.indexOf('gzip:') === 0
      ? Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(value.slice(5)))).getDataAsString('UTF-8')
      : value;
    return JSON.parse(json);
  } catch (error) {
    console.warn('Cache read dilewati:', error && error.message ? error.message : error);
    return null;
  }
}

function putCachedJson_(key, value, seconds) {
  if (typeof CacheService === 'undefined') return;
  try {
    let json = JSON.stringify(value);
    const blob = Utilities.newBlob(json, 'application/json');
    let bytes = blob.getBytes().length;
    // Compress large responses so the same ledger need not be reread for every
    // page visit. Keep one atomic item; an evicted/corrupt item is a cache miss.
    if (bytes >= 95000) {
      json = 'gzip:' + Utilities.base64Encode(Utilities.gzip(blob).getBytes());
      bytes = json.length; // Base64 and prefix are ASCII.
    }
    if (bytes < 95000) {
      CacheService.getScriptCache().put('km-app-v8:' + key, json, seconds || 45);
    }
  } catch (error) {
    console.warn('Cache write dilewati:', error && error.message ? error.message : error);
  }
}

function invalidateAppCaches_() {
  KM_REQUEST_STATUS_OPTIONS_ = null;
  KM_REQUEST_CONFIG_PARAMETERS_ = null;
  KM_REQUEST_CONFIG_COVERAGE_ = null;
  KM_REQUEST_UNIT_COVERAGE_ = null;
  KM_REQUEST_EFFECTIVE_COVERAGE_ = null;
  if (typeof CacheService === 'undefined') return;
  try {
    CacheService.getScriptCache().removeAll([
      'km-app-v8:monitoring', 'km-app-v8:master-unit',
      'km-app-v8:master-unit-options', 'km-app-v8:master-dropdowns',
      'km-app-v8:config-lists', 'km-app-v8:config-parameters',
      'km-app-v8:status-3l-options', 'km-app-v8:analytics-revision',
      'km-app-v8:admin-health-deep-v23',
      'km-app-v8:admin-health-deep-v24'
    ]);
  } catch (error) {
    console.warn('Cache invalidation dilewati:', error && error.message ? error.message : error);
  }
}

function invalidateUserDirectoryCache_() {
  KM_REQUEST_USER_DIRECTORY_ = null;
  if (typeof CacheService === 'undefined') return;
  try {
    CacheService.getScriptCache().remove('km-app-v8:users');
  } catch (error) {
    console.warn('Cache user tidak dapat dihapus:', error && error.message ? error.message : error);
  }
}

function invalidateRoleDirectoryCache_() {
  KM_REQUEST_ROLE_DIRECTORY_ = null;
  KM_REQUEST_ACTIVE_ROLE_NAMES_ = null;
  if (typeof CacheService === 'undefined') return;
  try {
    CacheService.getScriptCache().remove('km-app-v8:roles');
  } catch (error) {
    console.warn('Cache role tidak dapat dihapus:', error && error.message ? error.message : error);
  }
}

function invalidateSecurityCaches_() {
  // Base data cache bersifat scope-neutral; setiap response tetap difilter
  // menggunakan actor yang baru divalidasi.
  invalidateUserDirectoryCache_();
  invalidateRoleDirectoryCache_();
}

function ensureHeaders_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    if (headers.length > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
    sheet
      .getRange(1, 1, 1, headers.length)
      .setValues([headers]);
    return;
  }

  const existing = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0]
    .map(value => String(value).trim());

  const normalized = existing.map(normalizeHeaderKey_);
  headers.forEach(header => {
    const indexes = normalized.reduce((list, key, index) => {
      if (key === normalizeHeaderKey_(header)) list.push(index);
      return list;
    }, []);
    if (indexes.length > 1) throw new Error('Header ganda ' + header + ' di ' + sheet.getName() + '.');
    if (indexes.length === 1 && existing[indexes[0]] !== header) sheet.getRange(1, indexes[0] + 1).setValue(header);
  });
  const missing = headers.filter(header => !normalized.includes(normalizeHeaderKey_(header)));

  if (missing.length) {
    const requiredColumns = sheet.getLastColumn() + missing.length;
    if (requiredColumns > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
    sheet
      .getRange(
        1,
        sheet.getLastColumn() + 1,
        1,
        missing.length
      )
      .setValues([missing]);
  }
}

function ensureSheetCapacity_(sheet, extraRows) {
  const needed = sheet.getLastRow() + extraRows;

  if (needed > sheet.getMaxRows()) {
    sheet.insertRowsAfter(
      sheet.getMaxRows(),
      needed - sheet.getMaxRows()
    );
  }
}

/**
 * Menentukan row tujuan untuk tabel append-only berdasarkan kolom key canonical.
 *
 * Row tanpa key dianggap tidak pernah menjadi record yang valid. Row tersebut
 * boleh dipakai ulang selama tidak memiliki formula/note, sehingga sisa update
 * sesi yang salah alamat tidak membuat nomor row berikutnya terus meloncat.
 * Pemanggil wajib sudah memegang ScriptLock karena pencarian + penulisan harus
 * menjadi satu transaksi.
 */
function findReusableCanonicalRowUnsafe_(sheet, keyHeader) {
  const map = getHeaderMap_(sheet);
  const keyColumn = map[keyHeader];
  if (!keyColumn) throw new Error('Header key ' + keyHeader + ' tidak ditemukan di ' + sheet.getName() + '.');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;

  const keys = sheet.getRange(2, keyColumn, lastRow - 1, 1).getDisplayValues();
  for (let offset = 0; offset < keys.length; offset++) {
    if (cleanText_(keys[offset][0])) continue;
    const rowNumber = offset + 2;
    const candidate = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn());
    const hasFormula = candidate.getFormulas()[0].some(Boolean);
    const hasNote = candidate.getNotes()[0].some(Boolean);
    if (!hasFormula && !hasNote) return rowNumber;
  }
  return lastRow + 1;
}

function appendCanonicalRowUnsafe_(sheet, row, keyHeader) {
  const rowNumber = findReusableCanonicalRowUnsafe_(sheet, keyHeader);
  if (rowNumber > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), rowNumber - sheet.getMaxRows());
  }
  sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).setValues([row]);
  return rowNumber;
}

function getHeaderMap_(sheet) {
  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];

  const map = {};

  headers.forEach((header, index) => {
    const key = String(header || '').trim();

    if (key) {
      map[key] = index + 1;
    }
  });

  return map;
}

function setValidation_(sheet, column, rowCount, sourceRange) {
  if (!column) {
    return;
  }

  const rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(false)
    .build();

  try {
    sheet
      .getRange(2, column, rowCount, 1)
      .setDataValidation(rule);
  } catch (error) {
    if (isTypedColumnError_(error)) {
      const message = String(error && error.message ? error.message : error);
      console.warn('Data validation dilewati untuk typed column: ' + message);
      return;
    }

    throw error;
  }
}

function isTypedColumnError_(error) {
  const message = String(error && error.message ? error.message : error);
  return /typed\s+columns?/i.test(message);
}


function trySetNumberFormat_(range, format) {
  try {
    range.setNumberFormat(format);
    return true;
  } catch (error) {
    if (isTypedColumnError_(error)) {
      const message = String(error && error.message ? error.message : error);
      console.warn('Format tanggal dilewati untuk typed column: ' + message);
      return false;
    }

    throw error;
  }
}


function runNonCritical_(label, task) {
  try {
    task();
    return true;
  } catch (error) {
    if (isTypedColumnError_(error)) {
      console.warn(
        'Proses presentasi dilewati karena typed column (' + label + '): ' +
        String(error && error.message ? error.message : error)
      );
      return false;
    }

    console.error(
      'Proses non-kritis gagal (' + label + '):',
      error && error.stack ? error.stack : error
    );
    return false;
  }
}

function setRowValue_(row, map, header, value) {
  if (!map[header]) {
    return;
  }

  row[map[header] - 1] = value;
}

function getRowValue_(row, map, header) {
  if (!map[header]) {
    return '';
  }

  return row[map[header] - 1];
}

function setCellByHeader_(sheet, map, row, header, value) {
  if (!map[header]) {
    return;
  }

  sheet
    .getRange(row, map[header])
    .setValue(value);
}

function setCellByHeaderIfBlank_(sheet, map, row, header, value) {
  if (!map[header]) {
    return;
  }

  const cell = sheet.getRange(row, map[header]);

  if (cell.getValue() === '') {
    cell.setValue(value);
  }
}

function getCellByHeader_(sheet, map, row, header) {
  if (!map[header]) {
    return '';
  }

  return sheet
    .getRange(row, map[header])
    .getValue();
}

function uniqueClean_(values) {
  return [
    ...new Set(
      values
        .map(value => String(value || '').trim())
        .filter(Boolean)
    )
  ].sort((a, b) => a.localeCompare(b, 'id'));
}

function cleanText_(value) {
  return String(value || '').trim();
}

function normalizeLookupText_(value) {
  return cleanText_(value)
    .replace(/\s+/g, ' ')
    .toUpperCase();
}


function safeSheetText_(value) {
  const text = cleanText_(value);

  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function normalizeNopol_(value) {
  return String(value || '')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '')
    .trim();
}

// Regex untuk TextFinder yang tetap menganggap AB-8269-B, AB 8269 B, dan
// AB8269B sebagai nomor polisi yang sama. Pencarian terjadi di sisi Sheets,
// sehingga dashboard tidak perlu menarik seluruh kolom Nopol ke runtime.
function normalizedNopolTextFinderPattern_(value) {
  const target = normalizeNopol_(value);
  return target ? '^' + target.split('').join('[^A-Za-z0-9]*') + '$' : '';
}

function formatNopol_(value) {
  const compact = normalizeNopol_(value);
  const match = compact.match(/^([A-Z]{1,2})(\d{1,4})([A-Z]{0,3})$/);
  if (!match) return String(value || '').toUpperCase().trim();
  return match[1] + '-' + match[2] + (match[3] ? '-' + match[3] : '');
}

function isCanonicalNopol_(value) {
  const text = String(value || '').toUpperCase().trim();
  return /^[A-Z]{1,2}-\d{1,4}(?:-[A-Z]{1,3})?$/.test(text) && text === formatNopol_(text);
}

function normalizeDisplayNopol_(value) {
  return formatNopol_(value);
}

function parseDateInput_(value) {
  if (value === '' || value === null || value === undefined) {
    return '';
  }

  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';
    return createValidDate_(value.getFullYear(), value.getMonth() + 1, value.getDate());
  }

  const str = String(value).trim();

  if (!str) {
    return '';
  }

  let match = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);

  if (match) {
    return createValidDate_(
      Number(match[1]),
      Number(match[2]),
      Number(match[3])
    );
  }

  match = str.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$/);

  if (match) {
    return createValidDate_(
      Number(match[3]),
      Number(match[2]),
      Number(match[1])
    );
  }

  match = str.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);
  if (match) {
    return createValidDate_(Number(match[3]), parseIndonesianMonth_(match[2]), Number(match[1]));
  }

  if (/^\d+(\.\d+)?$/.test(str)) {
    const serial = Number(str);

    if (serial > 20000 && serial < 80000) {
      const excelDate = new Date(Date.UTC(1899, 11, 30) + Math.floor(serial) * 86400000);

      return createValidDate_(
        excelDate.getUTCFullYear(),
        excelDate.getUTCMonth() + 1,
        excelDate.getUTCDate()
      );
    }
  }

  return '';
}

function createValidDate_(year, month, day) {
  if (!Number.isInteger(year) || !Number.isInteger(month) ||
      !Number.isInteger(day) || year < 1900 || year > 2100) {
    return '';
  }

  // Pukul 12 lokal mencegah pergeseran hari saat objek melewati serialisasi/timezone.
  const date = new Date(year, month - 1, day, 12, 0, 0, 0);

  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    return '';
  }

  return date;
}

function dateOnlyKey_(value) {
  const date = value instanceof Date ? value : parseDateInput_(value);
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatDateTime_(value) {
  const date = parseDateTimeInput_(value);
  if (!date) return displayValue_(value);
  const parts = dateTimeParts_(date);
  return formatIndonesianDateParts_(parts.year, parts.month, parts.day) +
    ', ' + pad2_(parts.hour) + '.' + pad2_(parts.minute);
}

function formatDateOnly_(value) {
  const date = value instanceof Date ? value : parseDateInput_(value);
  if (!(date instanceof Date) || isNaN(date.getTime())) return displayValue_(value);
  const parts = dateTimeParts_(date);
  return formatIndonesianDateParts_(parts.year, parts.month, parts.day);
}

function parseDateTimeInput_(value) {
  if (value instanceof Date) return isNaN(value.getTime()) ? '' : value;
  const text = cleanText_(value);
  if (!text) return '';
  // Sheet serials preserve fractional time; JSON cache dates retain their UTC offset.
  if (/^\d+(?:\.\d+)?$/.test(text) && Number(text) > 20000 && Number(text) < 80000) {
    const wall = new Date(Date.UTC(1899, 11, 30) + Math.round(Number(text) * 86400000));
    return createValidDateTime_(wall.getUTCFullYear(), wall.getUTCMonth() + 1,
      wall.getUTCDate(), wall.getUTCHours(), wall.getUTCMinutes(), wall.getUTCSeconds());
  }
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,3})?(?:Z|[+-]\d{2}:?\d{2})$/.test(text)) {
    const parts = text.slice(0, 19).split(/[-T:]/).map(Number);
    if (!createValidDateTime_(parts[0], parts[1], parts[2], parts[3], parts[4], parts[5])) return '';
    const date = new Date(text);
    return isNaN(date.getTime()) ? '' : date;
  }

  let match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (match) {
    return createValidDateTime_(Number(match[1]), Number(match[2]), Number(match[3]),
      Number(match[4] || 0), Number(match[5] || 0), Number(match[6] || 0));
  }

  match = text.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (match) {
    return createValidDateTime_(Number(match[3]), Number(match[2]), Number(match[1]),
      Number(match[4] || 0), Number(match[5] || 0), Number(match[6] || 0));
  }

  match = text.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})(?:,?\s+(\d{1,2})[.:](\d{2})(?:[.:](\d{2}))?)?$/);
  if (match) {
    return createValidDateTime_(Number(match[3]), parseIndonesianMonth_(match[2]), Number(match[1]),
      Number(match[4] || 0), Number(match[5] || 0), Number(match[6] || 0));
  }

  return '';
}

function createValidDateTime_(year, month, day, hour, minute, second) {
  if (![year, month, day, hour, minute, second].every(Number.isInteger) ||
      year < 1900 || year > 2100 || hour < 0 || hour > 23 ||
      minute < 0 || minute > 59 || second < 0 || second > 59) return '';
  const date = new Date(year, month - 1, day, hour, minute, second, 0);
  return date.getFullYear() === year && date.getMonth() === month - 1 &&
    date.getDate() === day && date.getHours() === hour && date.getMinutes() === minute &&
    date.getSeconds() === second ? date : '';
}

const INDONESIAN_MONTH_NAMES_ = Object.freeze([
  'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
]);

function parseIndonesianMonth_(value) {
  const key = cleanText_(value).toUpperCase();
  const aliases = {
    JAN: 1, JANUARI: 1, JANUARY: 1,
    FEB: 2, FEBRUARI: 2, FEBRUARY: 2,
    MAR: 3, MARET: 3, MARCH: 3,
    APR: 4, APRIL: 4,
    MEI: 5, MAY: 5,
    JUN: 6, JUNI: 6, JUNE: 6,
    JUL: 7, JULI: 7, JULY: 7,
    AGU: 8, AGS: 8, AGUSTUS: 8, AUG: 8, AUGUST: 8,
    SEP: 9, SEPT: 9, SEPTEMBER: 9,
    OKT: 10, OKTOBER: 10, OCT: 10, OCTOBER: 10,
    NOV: 11, NOVEMBER: 11,
    DES: 12, DESEMBER: 12, DEC: 12, DECEMBER: 12
  };
  return aliases[key] || 0;
}

function dateTimeParts_(date) {
  const raw = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd-HH-mm-ss').split('-').map(Number);
  return { year: raw[0], month: raw[1], day: raw[2], hour: raw[3], minute: raw[4], second: raw[5] };
}

function pad2_(value) {
  return String(value).padStart(2, '0');
}

function formatIndonesianDateParts_(year, month, day) {
  return pad2_(day) + ' ' + (INDONESIAN_MONTH_NAMES_[month - 1] || '') + ' ' + year;
}

function formatMasterUnitDetailValue_(header, value) {
  const key = cleanText_(header).toUpperCase().replace(/\s+/g, ' ');
  if (MASTER_UNIT_FIELD_ALIASES.status3LUpdatedBy.some(alias =>
      normalizeHeaderKey_(alias) === normalizeHeaderKey_(header))) {
    return status3LActorDisplay_(displayValue_(value));
  }
  const isDateOnly = /(^| )(TANGGAL|TGL|DATE)( |$)/.test(key);
  const isDateTime = key === 'TIMESTAMP' || / AT$/.test(key) || / UNTIL$/.test(key);
  if (isDateOnly) return formatDateOnly_(value);
  if (isDateTime || value instanceof Date) return formatDateTime_(value);
  return displayValue_(value);
}

function displayValue_(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value);
}

function getCurrentUserEmail_() {
  // Hanya untuk konteks Spreadsheet / audit fallback. Web App memakai custom session token.
  return getSheetUserEmail_();
}

/**
 * Mengecilkan grid Sheet canonical agar database tidak dipenuhi ribuan row/column kosong.
 * Hanya dijalankan saat Setup/Migrasi, bukan pada request dashboard biasa.
 */
function compactDatabaseGrid_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const specs = [
    [KM_APP.SHEETS.CONFIG, 6, 20, 20],
    [KM_APP.SHEETS.MASTER_UNIT, null, 100, 200],
    [KM_APP.SHEETS.USERS, KM_APP.USER_HEADERS.length, 20, 20],
    [KM_APP.SHEETS.ROLES, KM_ROLE_HEADERS.length, 20, 20],
    [KM_APP.SHEETS.SESSIONS, KM_AUTH.SESSION_HEADERS.length, 50, 50],
    [KM_APP.SHEETS.AUTH_LOG, KM_AUTH.AUTH_LOG_HEADERS.length, 100, 100],
    [KM_APP.SHEETS.LOG, KM_APP.LOG_HEADERS.length, 100, 100],
    [UNIT_KM_LOG_SHEET, UNIT_KM_LOG_HEADERS.length, 100, 100],
    [ANALYTICS_CALIBRATION_SHEET, ANALYTICS_CALIBRATION_HEADERS.length, 100, 101],
    [UNIT_STATUS_LOG_SHEET, UNIT_STATUS_LOG_HEADERS.length, 100, 100],
    [UNIT_STATUS_3L_LOG_SHEET, UNIT_STATUS_3L_LOG_HEADERS.length, 100, 100],
    [MASTER_UNIT_MANAGEMENT.MOVEMENT_SHEET, MASTER_UNIT_MANAGEMENT.MOVEMENT_HEADERS.length, 100, 101],
    [MASTER_UNIT_MANAGEMENT.REQUEST_SHEET, MASTER_UNIT_MANAGEMENT.REQUEST_HEADERS.length, 100, 101]
  ];

  const changes = [];
  specs.forEach(spec => {
    const sheet = ss.getSheetByName(spec[0]);
    if (!sheet) return;
    changes.push(trimSheetGrid_(sheet, spec[1], spec[2], spec[3]));
  });
  return {
    sheets: changes,
    removedGridRows: changes.reduce((total, item) => total + item.removedRows, 0),
    removedGridColumns: changes.reduce((total, item) => total + item.removedColumns, 0)
  };
}

function trimSheetGrid_(sheet, exactColumns, rowBuffer, minRows) {
  const beforeRows = sheet.getMaxRows();
  const beforeColumns = sheet.getMaxColumns();
  const lastRow = Math.max(1, sheet.getLastRow());
  const targetRows = Math.max(Number(minRows) || 20, lastRow + Math.max(0, Number(rowBuffer) || 0));
  if (sheet.getMaxRows() > targetRows) {
    sheet.deleteRows(targetRows + 1, sheet.getMaxRows() - targetRows);
  } else if (sheet.getMaxRows() < targetRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), targetRows - sheet.getMaxRows());
  }

  const targetColumns = Math.max(1, Number(exactColumns) || 0, sheet.getLastColumn());
  if (sheet.getMaxColumns() > targetColumns) {
    sheet.deleteColumns(targetColumns + 1, sheet.getMaxColumns() - targetColumns);
  } else if (sheet.getMaxColumns() < targetColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), targetColumns - sheet.getMaxColumns());
  }
  return {
    sheet: sheet.getName(),
    beforeRows: beforeRows,
    afterRows: sheet.getMaxRows(),
    beforeColumns: beforeColumns,
    afterColumns: sheet.getMaxColumns(),
    removedRows: Math.max(0, beforeRows - sheet.getMaxRows()),
    removedColumns: Math.max(0, beforeColumns - sheet.getMaxColumns())
  };
}

/** Migrasi eksplisit melalui menu Setup; tidak berjalan saat login/update unit. */


// ============================================================================
// 04. UNIT OPERATIONS & PHOTO HISTORY
// ============================================================================
/** Master Unit update workflow and unit-history query endpoints. */
function auditMasterUnitKmDatabase(sessionToken) {
  requirePermission_('ADMIN_SETUP', '', { token: sessionToken });
  return auditMasterUnitKmDatabase_();
}

function auditMasterUnitKmDatabase_(auditContext) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  const required = { nopol: 'NOMOR POLISI', operatingPoint: 'OPERATING POINT', toh: 'NAMA TOH/WOH', ...UNIT_KM_FIELDS, ...UNIT_MASTER_DISPLAY_FIELDS };
  const result = { sheetFound: Boolean(sheet), missingColumns: [], duplicateHeaders: [], duplicateNopol: [], invalidNopolFormatRows: [],
    missingNopolRows: [], missingAreaRows: [], missingTohRows: [], invalidStatusRows: [], invalidKmRows: [], incompleteKmRows: [],
    missingKmDateRows: [], invalidKmDateRows: [], invalidCalibrationRows: [], calibrationValidationCount: 0,
    status3LAttributionMismatchRows: [], status3LAttributionUnresolvedRows: [],
    status3LUnsupportedDoneRows: [], status3LStateMismatchRows: [], status3LStateUnresolvedRows: [],
    workflowColumnsReady: false, statusValidationReady: false, historyHidden: false,
    protectionMode: 'BACKEND_RBAC_ONLY',
    editTriggerReady: isKalibrasiKmEditTriggerReady_() };
  const values = sheet && sheet.getLastRow() ? sheet.getDataRange().getValues() : [[]];
  const headers = values[0].map(String);
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  Object.keys(required).forEach(field => {
    if (indexes[field] < 0) result.missingColumns.push(required[field]);
    const count = headers.filter(header => MASTER_UNIT_FIELD_ALIASES[field].some(alias => normalizeHeaderKey_(header) === normalizeHeaderKey_(alias))).length;
    if (count > 1) result.duplicateHeaders.push(required[field]);
  });
  const seen = new Set();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  values.slice(1).forEach((row, index) => {
    if (!row.some(value => value !== '' && value != null)) return;
    const number = index + 2;
    const get = field => indexes[field] < 0 ? '' : row[indexes[field]];
    const key = normalizeNopol_(get('nopol'));
    if (!key) result.missingNopolRows.push(number);
    else if (seen.has(key)) result.duplicateNopol.push({ row: number, nopol: String(get('nopol')) });
    if (key && !isCanonicalNopol_(get('nopol'))) result.invalidNopolFormatRows.push({ row: number, nopol: String(get('nopol')) });
    seen.add(key);
    if (!cleanText_(get('operatingPoint'))) result.missingAreaRows.push(number);
    if (!cleanText_(get('toh'))) result.missingTohRows.push(number);
    const statusUnitValue = displayValue_(get('statusUnit'));
    const calibrationSequence = parseCalibrationNumber_(get('calibrationNumber'));
    const status3LValue = displayValue_(get('status3L'));
    const initialBlankStatus = !cleanText_(statusUnitValue) && !calibrationSequence && !cleanText_(status3LValue);
    if (!initialBlankStatus && !UNIT_OPERATIONAL_STATUSES.some(status =>
        normalizeLookupText_(status) === normalizeLookupText_(statusUnitValue))) {
      result.invalidStatusRows.push(number);
    }
    const readings = ['kmOdometer'].map(get);
    if (readings.some(value => value === '' || value == null)) result.incompleteKmRows.push(number);
    if (readings.some(value => value !== '' && value != null && parseUnitKm_(value) === null)) result.invalidKmRows.push(number);
    if (!get('kmDate')) result.missingKmDateRows.push(number);
    else if (!unitKmDate_(get('kmDate')) || unitKmDate_(get('kmDate')) > today) result.invalidKmDateRows.push(number);
    if (get('calibrationNumber') !== '' && get('calibrationNumber') != null &&
        !parseCalibrationNumber_(get('calibrationNumber'))) result.invalidCalibrationRows.push(number);
  });
  result.workflowColumnsReady = indexes.kmDate >= 0 &&
    indexes.kmUpdatedAt === indexes.kmDate + 1 &&
    indexes.kmUpdatedBy === indexes.kmUpdatedAt + 1 &&
    indexes.calibrationNumber === indexes.kmUpdatedBy + 1 &&
    indexes.status3L === indexes.calibrationNumber + 1 &&
    indexes.status3LUpdatedAt === indexes.status3L + 1 &&
    indexes.status3LUpdatedBy === indexes.status3LUpdatedAt + 1 &&
    indexes.status3LNote === indexes.status3LUpdatedBy + 1;
  if (sheet && indexes.calibrationNumber >= 0 && sheet.getMaxRows() > 0) {
    result.calibrationValidationCount = sheet
      .getRange(1, indexes.calibrationNumber + 1, sheet.getMaxRows(), 1)
      .getDataValidations()
      .reduce((total, row) => total + (row[0] ? 1 : 0), 0);
  }
  const log = getUnitKmHistorySheet_(ss);
  result.historySheetReady = historyHeadersReady_(log, UNIT_KM_LOG_HEADERS);
  result.historyHidden = Boolean(log && log.isSheetHidden());
  result.statusValidationReady = isMasterUnitStatusValidationReady_(sheet, indexes.statusUnit);
  const statusLog = ss.getSheetByName(UNIT_STATUS_LOG_SHEET);
  result.statusHistorySheetReady = historyHeadersReady_(statusLog, UNIT_STATUS_LOG_HEADERS);
  const status3LLog = ss.getSheetByName(UNIT_STATUS_3L_LOG_SHEET);
  result.status3LHistorySheetReady = historyHeadersReady_(status3LLog, UNIT_STATUS_3L_LOG_HEADERS);
  if (result.status3LHistorySheetReady && sheet && sheet.getLastRow() > 1 && indexes.status3LUpdatedBy >= 0) {
    const auditMap = buildStatus3LAuditEvidenceMap_(status3LLog);
    const activityMap = buildStatus3LActivityEvidenceMap_(ss.getSheetByName(KM_APP.SHEETS.LOG));
    values.slice(1).forEach((row, offset) => {
      const status3L = indexes.status3L >= 0 ? displayValue_(row[indexes.status3L]) : '';
      if (!status3L) return;
      const evidence = status3LAttributionEvidenceForRow_(auditMap, activityMap, row, indexes);
      if (!evidence || !evidence.resolvedActor) {
        if (result.status3LAttributionUnresolvedRows.length < 100) result.status3LAttributionUnresolvedRows.push({
          row: offset + 2,
          nopol: displayValue_(row[indexes.nopol]),
          calibrationNumber: parseCalibrationNumber_(row[indexes.calibrationNumber]),
          status3L: status3L,
          currentUpdatedBy: status3LActorDisplay_(displayValue_(row[indexes.status3LUpdatedBy]))
        });
        return;
      }
      const currentUpdatedBy = status3LActorDisplay_(displayValue_(row[indexes.status3LUpdatedBy]));
      if (normalizeLookupText_(currentUpdatedBy) !== normalizeLookupText_(evidence.updatedBy)) {
        if (result.status3LAttributionMismatchRows.length < 100) result.status3LAttributionMismatchRows.push({
          row: offset + 2,
          nopol: displayValue_(row[indexes.nopol]),
          calibrationNumber: parseCalibrationNumber_(row[indexes.calibrationNumber]),
          status3L: status3L,
          currentUpdatedBy: currentUpdatedBy,
          expectedUpdatedBy: evidence.updatedBy,
          source: evidence.source,
          reason: evidence.reason
        });
      }
    });
  }
  if (sheet && sheet.getLastRow() > 1 && indexes.nopol >= 0 && indexes.calibrationNumber >= 0 && indexes.status3L >= 0 && indexes.statusUnit >= 0) {
    const trustedStatus = auditContext && auditContext.trustedStatus ||
      buildTrustedStatus3LStateIndex_(status3LLog);
    // Reuse the status index already loaded above instead of rereading the full
    // STATUS_3L_HISTORY ledger through buildCompletedCalibrationKeySet_.
    const completedKeys = new Set();
    Object.keys(trustedStatus).forEach(key => {
      if (trustedStatus[key] && trustedStatus[key].done) completedKeys.add(key);
    });
    const trustedFinals = auditContext && auditContext.trustedFinals ||
      buildTrustedFinalSnapshotIndex_(trustedStatus);
    Object.keys(trustedFinals).forEach(key => completedKeys.add(key));
    values.slice(1).forEach((row, offset) => {
      const nopol = normalizeDisplayNopol_(row[indexes.nopol]);
      const sequence = parseCalibrationNumber_(row[indexes.calibrationNumber]);
      if (!nopol || !sequence) return;
      const key = calibrationCompletionKey_(nopol, sequence);
      const current3L = displayValue_(row[indexes.status3L]);
      const currentUnit = displayValue_(row[indexes.statusUnit]);
      const bucket = trustedStatus[key] || null;
      const latest = bucket && bucket.latest || null;
      const completionTrusted = completedKeys.has(key);
      if (normalizeLookupText_(current3L) === 'DONE' && !completionTrusted) {
        if (result.status3LUnsupportedDoneRows.length < 100) result.status3LUnsupportedDoneRows.push({ row: offset + 2, nopol, calibrationNumber: sequence, statusUnit: currentUnit, status3L: current3L });
      }
      if (latest) {
        const expected3L = displayValue_(latest.status3L);
        const onDelivery = normalizeLookupText_(currentUnit) === normalizeLookupText_('On Delivery');
        const expectedCurrent3L = onDelivery ? '' : expected3L;
        const expectedUnit = normalizeLookupText_(expectedCurrent3L) === 'DONE' ? 'Ready to Depart' : (onDelivery ? 'On Delivery' : 'Standby');
        if (normalizeLookupText_(current3L) !== normalizeLookupText_(expectedCurrent3L) ||
            (normalizeLookupText_(currentUnit) === normalizeLookupText_('Ready to Depart') && normalizeLookupText_(expectedUnit) !== normalizeLookupText_('Ready to Depart')) ||
            (normalizeLookupText_(expectedCurrent3L) === 'DONE' && normalizeLookupText_(currentUnit) !== normalizeLookupText_('Ready to Depart'))) {
          if (result.status3LStateMismatchRows.length < 100) result.status3LStateMismatchRows.push({ row: offset + 2, nopol, calibrationNumber: sequence, currentStatus3L: current3L, expectedStatus3L: expectedCurrent3L, currentStatusUnit: currentUnit, expectedStatusUnit: expectedUnit, evidenceSource: latest.source, evidenceReason: latest.reason });
        }
      } else if (cleanText_(current3L) && !completionTrusted) {
        if (result.status3LStateUnresolvedRows.length < 100) result.status3LStateUnresolvedRows.push({ row: offset + 2, nopol, calibrationNumber: sequence, statusUnit: currentUnit, status3L: current3L });
      }
    });
  }
  result.recommendation = 'Gunakan Full Safe Repair untuk standardisasi schema, RBAC, workflow evidence-first, photo integrity, history context, trigger, dan struktur. MASTER_UNIT tidak boleh dipakai sebagai bukti independen untuk memvalidasi dirinya sendiri.';
  console.log(JSON.stringify({
    audit: 'MASTER_UNIT', sheetFound: result.sheetFound,
    missingColumns: result.missingColumns.length,
    duplicateNopol: result.duplicateNopol.length,
    invalidStatusRows: result.invalidStatusRows.length,
    attributionMismatch: result.status3LAttributionMismatchRows.length,
    unresolvedState: result.status3LStateUnresolvedRows.length
  }));
  return result;
}

function parseUnitKm_(value) {
  if (value === '' || value == null) return null;
  // Odometer adalah bilangan bulat: hanya digit 0-9, tanpa tanda baca,
  // pemisah ribuan, desimal, notasi eksponen, atau simbol lain.
  const text = String(value).trim();
  if (!/^\d+$/.test(text)) return null;
  const number = Number(text);
  return Number.isSafeInteger(number) && number >= 0 && number <= 999999999 ? number : null;
}

function nextCalibrationNumber_(historySheet, historyMap, nopol, currentValue) {
  let maximum = parseCalibrationNumber_(currentValue);
  let legacyUnnumberedMatches = 0;
  if (!historySheet || historySheet.getLastRow() < 2) return Math.max(1, maximum + 1);
  const rowCount = historySheet.getLastRow() - 1;
  const nopolColumn = historyMap['NOMOR POLISI'];
  const sequenceColumn = historyMap['KALIBRASI KE'];
  const firstColumn = Math.min(nopolColumn, sequenceColumn);
  const historyValues = historySheet
    .getRange(2, firstColumn, rowCount, Math.abs(sequenceColumn - nopolColumn) + 1)
    .getValues();
  const nopolOffset = nopolColumn - firstColumn;
  const sequenceOffset = sequenceColumn - firstColumn;
  const target = normalizeNopol_(nopol);
  historyValues.forEach(row => {
    if (normalizeNopol_(row[nopolOffset]) !== target) return;
    const sequence = parseCalibrationNumber_(row[sequenceOffset]);
    if (sequence) maximum = Math.max(maximum, sequence);
    else legacyUnnumberedMatches++;
  });
  // Snapshot final dapat menghasilkan lebih dari satu record untuk sequence
  // yang sama. Hanya row legacy tanpa nomor yang dipakai sebagai fallback.
  return Math.max(maximum, legacyUnnumberedMatches) + 1;
}

function unitKmDate_(value) {
  return dateOnlyKey_(value);
}

function unitKmSnapshot_(row, indexes) {
  const get = field => indexes[field] >= 0 ? row[indexes[field]] : '';
  const updatedAt = get('kmUpdatedAt');
  const status3LUpdatedAt = get('status3LUpdatedAt');
  const snapshot = {
    statusUnit: String(get('statusUnit') || ''),
    status3L: String(get('status3L') || ''),
    status3LNote: String(get('status3LNote') || ''),
    status3LUpdatedAt: status3LUpdatedAt instanceof Date
      ? status3LUpdatedAt.toISOString() : String(status3LUpdatedAt || ''),
    status3LUpdatedBy: normalizeDoneStatus3LActor_(get('status3L'), String(get('status3LUpdatedBy') || '')),
    kmOdometer: parseUnitKm_(get('kmOdometer')),
    kmDate: unitKmDate_(get('kmDate')),
    kmUpdatedAt: updatedAt instanceof Date ? updatedAt.toISOString() : String(updatedAt || ''),
    kmUpdatedBy: String(get('kmUpdatedBy') || ''),
    kmPhotoId: String(get('kmPhotoId') || ''),
    kmPhotoAddedBy: String(get('kmPhotoAddedBy') || ''),
    kmPhotoAddedAt: get('kmPhotoAddedAt') instanceof Date ? get('kmPhotoAddedAt').toISOString() : String(get('kmPhotoAddedAt') || ''),
    calibrationNumber: parseCalibrationNumber_(get('calibrationNumber'))
  };
  // Sertakan nilai mentah agar perubahan langsung di Sheet juga terdeteksi.
  snapshot.kmRevision = JSON.stringify(
    Object.keys(UNIT_KM_FIELDS).concat(Object.keys(UNIT_MASTER_DISPLAY_FIELDS)).map(get)
  );
  return snapshot;
}

/**
 * Status 3L baru boleh diproses setelah ada bukti pembaruan odometer yang
 * lengkap. Pemeriksaan ini digunakan untuk payload UI dan divalidasi ulang
 * pada endpoint update agar tidak dapat dilewati dari browser console.
 */
function unitStatus3LEligibility_(snapshot) {
  const data = snapshot || {};
  const missing = [];
  if (parseUnitKm_(data.kmOdometer) === null) missing.push('KM odometer');
  if (!unitKmDate_(data.kmDate)) missing.push('tanggal KM');
  if (!cleanText_(data.kmPhotoId)) missing.push('foto odometer');
  if (!parseCalibrationNumber_(data.calibrationNumber)) missing.push('nomor kalibrasi');
  const operationalStatus = normalizeLookupText_(data.statusUnit);
  const status3L = normalizeLookupText_(data.status3L);
  let operationalReason = '';
  if (operationalStatus === normalizeLookupText_('On Delivery')) {
    operationalReason = 'Unit sedang On Delivery. Status 3L sengaja dikosongkan dan kalibrasi hanya dilakukan saat Standby.';
  } else if (operationalStatus && operationalStatus !== normalizeLookupText_('Standby') &&
      !(operationalStatus === normalizeLookupText_('Ready to Depart') && status3L === 'DONE')) {
    operationalReason = 'Status Unit harus Standby sebelum proses 3L dilakukan.';
  }
  return {
    dataComplete: missing.length === 0,
    eligible: missing.length === 0 && !operationalReason,
    reason: operationalReason || (missing.length
      ? 'Update Unit / Perbaiki Database untuk melengkapi ' + missing.join(', ') + ' terlebih dahulu.'
      : '')
  };
}


function buildMasterUnitClientPatch_(row, indexes) {
  const kmSnapshot = unitKmSnapshot_(row, indexes);
  const status3LEligibility = unitStatus3LEligibility_(kmSnapshot);
  return {
    nopol: normalizeDisplayNopol_(masterUnitValue_(row, indexes.nopol)),
    operatingPoint: masterUnitValue_(row, indexes.operatingPoint),
    toh: masterUnitValue_(row, indexes.toh),
    statusUnit: masterUnitValue_(row, indexes.statusUnit),
    status3L: masterUnitValue_(row, indexes.status3L),
    status3LUpdatedAt: kmSnapshot.status3LUpdatedAt,
    status3LUpdatedBy: kmSnapshot.status3LUpdatedBy,
    status3LNote: masterUnitValue_(row, indexes.status3LNote),
    kmOdometer: kmSnapshot.kmOdometer,
    kmDate: kmSnapshot.kmDate,
    kmUpdatedAt: kmSnapshot.kmUpdatedAt,
    calibrationNumber: kmSnapshot.calibrationNumber,
    kmRevision: kmSnapshot.kmRevision,
    kmPhotoId: kmSnapshot.kmPhotoId,
    calibrationDataComplete: status3LEligibility.dataComplete,
    status3LEligible: status3LEligibility.eligible,
    status3LEligibilityReason: status3LEligibility.reason
  };
}

function validateUnitKmUpdate_(payload, unit, today, validatedPhoto) {
  if (!unit) throw new Error('Nomor polisi tidak ditemukan di MASTER_UNIT.');
  if (unit.duplicate) throw new Error('Nomor polisi duplikat di MASTER_UNIT; perbaiki master terlebih dahulu.');
  if (typeof payload.kmRevision !== 'string' || payload.kmRevision !== unit.snapshot.kmRevision) {
    throw new Error('Data unit sudah berubah. Refresh Pengajuan Kalibrasi lalu ulangi update.');
  }

  // v24.1: Pengajuan Kalibrasi KM adalah event pembuka workflow.
  // Unit baru boleh memiliki STATUS UNIT kosong; saat pengajuan kalibrasi dibuat,
  // backend authoritative menetapkan STATUS UNIT menjadi Standby.
  const status = 'Standby';

  const kmDate = unitKmDate_(payload.kmDate);
  if (!kmDate || kmDate > today) throw new Error('Tanggal KM tidak valid atau melebihi hari ini.');
  if (unit.snapshot.kmDate && kmDate < unit.snapshot.kmDate) {
    throw new Error('Tanggal KM lebih lama dari data terakhir (' + formatDateOnly_(unit.snapshot.kmDate) + ').');
  }
  const kmFields = ['kmOdometer'];
  const requestedTargets = kmFields.slice();
  if (!requestedTargets.length) throw new Error('Pilih minimal satu kategori KM yang akan diperbarui.');
  const result = { statusUnit: status, kmDate: kmDate, kmTargets: requestedTargets };
  kmFields.forEach(field => {
    if (!requestedTargets.includes(field)) {
      result[field] = unit.snapshot[field];
      return;
    }
    result[field] = parseUnitKm_(payload[field]);
    if (result[field] === null) throw new Error(UNIT_KM_FIELDS[field] + ' wajib hanya berisi digit angka 0-9 tanpa koma, titik, garis miring, atau simbol lain.');
  });
  result.photo = validatedPhoto || validateOdometerPhoto_(payload.photo);
  result.kmPhotoId = '';
  result.kmPhotoAddedAt = '';
  result.kmPhotoAddedBy = '';
  result.note = String(payload.note || '').trim();
  if (result.note.length > 1000) throw new Error('Catatan maksimal 1000 karakter.');
  if (unit.snapshot.kmDate === kmDate && unit.snapshot.kmOdometer === result.kmOdometer) {
    throw new Error('Odometer dan tanggal sama dengan kalibrasi terakhir. Masukkan hasil pengukuran baru agar riwayat tidak terduplikasi.');
  }
  if (unit.snapshot.kmOdometer !== null && result.kmOdometer < unit.snapshot.kmOdometer && !result.note) {
    throw new Error('Odometer lebih kecil dari data terakhir. Isi catatan alasan koreksi atau penggantian odometer.');
  }
  return result;
}

function updateMasterUnit(payload, sessionToken) {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload) || typeof payload.updateOdometer !== 'boolean') {
    throw new Error('Pilih jenis pembaruan unit yang valid.');
  }
  if (!payload.updateOdometer) {
    if (payload.photo || (payload.kmOdometer !== undefined && payload.kmOdometer !== '') ||
        (payload.kmDate !== undefined && payload.kmDate !== '')) {
      throw new Error('Aktifkan pembaruan odometer untuk menyimpan KM atau foto.');
    }
    return updateMasterUnitStatus_(payload, sessionToken);
  }
  return updateMasterUnitKm_(payload, sessionToken);
}

function getMasterUnitUpdateTarget_(sheet, nopol) {
  if (!sheet || sheet.getLastRow() < 2) throw new Error('MASTER_UNIT belum tersedia.');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(String);
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  const requiredFields = ['nopol', 'operatingPoint']
    .concat(Object.keys(UNIT_KM_FIELDS), Object.keys(UNIT_MASTER_DISPLAY_FIELDS));
  const missing = [];
  requiredFields.forEach(field => {
    const aliases = MASTER_UNIT_FIELD_ALIASES[field] || [];
    const count = headers.filter(header => aliases.some(alias =>
      normalizeHeaderKey_(header) === normalizeHeaderKey_(alias))).length;
    if (count > 1) throw new Error('DATABASE_NOT_READY: Header ganda untuk ' + field + '.');
    if (indexes[field] < 0) missing.push(field);
  });
  if (missing.length || indexes.kmUpdatedAt !== indexes.kmDate + 1 ||
      indexes.kmUpdatedBy !== indexes.kmUpdatedAt + 1 ||
      indexes.calibrationNumber !== indexes.kmUpdatedBy + 1 ||
      indexes.status3L !== indexes.calibrationNumber + 1 ||
      indexes.status3LUpdatedAt !== indexes.status3L + 1 ||
      indexes.status3LUpdatedBy !== indexes.status3LUpdatedAt + 1 ||
      indexes.status3LNote !== indexes.status3LUpdatedBy + 1) {
    throw new Error('DATABASE_NOT_READY: Struktur MASTER_UNIT belum siap. ' +
      'SUPER_ADMIN perlu menjalankan Perbaiki Database sebelum update unit.');
  }

  const key = normalizeNopol_(nopol);
  if (!key) throw new Error('Nomor polisi tidak valid.');
  const rowNumber = findMasterUnitRowNumberByNopol_(sheet, indexes, nopol);
  if (!rowNumber) throw new Error('Nomor polisi tidak ditemukan.');
  const usedIndexes = Object.keys(indexes).map(field => indexes[field]).filter(index => index >= 0);
  const readColumnCount = Math.max(1, usedIndexes.length ? Math.max.apply(null, usedIndexes) + 1 : 1);
  const row = sheet.getRange(rowNumber, 1, 1, readColumnCount).getValues()[0];
  if (indexes.masterStatus >= 0 && normalizeLookupText_(row[indexes.masterStatus]) === 'INACTIVE') {
    throw new Error('MASTER_UNIT_INACTIVE: unit berstatus INACTIVE dan tidak dapat diperbarui dari Monitoring. Reactivate melalui Administration > Master Unit.');
  }
  return {
    headers: headers,
    indexes: indexes,
    rowNumber: rowNumber,
    readColumnCount: readColumnCount,
    row: row,
    nopol: String(row[indexes.nopol] || '')
  };
}

/** Batch only adjacent editable fields; never write across a custom column.
 * Capture formulas and literal text so failed commits can restore the old cells.
 */
function buildUnitFieldWrites_(sheet, rowNumber, indexes, values) {
  const cells = Object.keys(values).map(field => {
    const index = indexes[field];
    if (!Number.isInteger(index) || index < 0) throw new Error('Kolom update tidak tersedia: ' + field);
    return { column: index + 1, value: values[field] };
  }).sort((a, b) => a.column - b.column);
  const groups = [];
  cells.forEach(cell => {
    const group = groups[groups.length - 1];
    if (group && cell.column === group.column + group.values.length) group.values.push(cell.value);
    else {
      if (group && cell.column < group.column + group.values.length) throw new Error('Kolom update duplikat.');
      groups.push({ column: cell.column, values: [cell.value] });
    }
  });
  return groups.map(group => {
    const range = sheet.getRange(rowNumber, group.column, 1, group.values.length);
    const oldValues = range.getValues()[0];
    const formulas = range.getFormulas()[0];
    return { range: range, next: [group.values], old: [oldValues.map((value, index) =>
      formulas[index] || (typeof value === 'string' ? safeSheetText_(value) : value))] };
  });
}

function updateMasterUnitKm_(payload, sessionToken) {
  const actor = requirePermission_('UPDATE_UNIT', '', { token: sessionToken });
  // Validasi backend bersifat authoritative; UI hanya memberi feedback awal.
  // Cek sebelum validasi/upload agar request di luar periode berhenti cepat.
  assertCalibrationSubmissionOpen_();
  let photoFile = null;
  let committed = false;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
    const preparedPhoto = validateOdometerPhoto_(payload.photo);
    const preflight = getMasterUnitUpdateTarget_(sheet, payload.nopol);
    const preflightOperatingPoint = String(preflight.row[preflight.indexes.operatingPoint] || '');
    if (!canAccessOperatingPoint_(actor, preflightOperatingPoint)) {
      throw new Error('Unit berada di luar scope Operating Point akun Anda.');
    }
    try {
      validateUnitKmUpdate_(payload, { snapshot: unitKmSnapshot_(preflight.row, preflight.indexes) },
        dateOnlyKey_(new Date()), preparedPhoto);
    } catch (error) {
      return { success: false, errors: [{ message: error.message }] };
    }

    const waitingStatus3L = resolveStatus3L_('WAITING');
    // Drive I/O sengaja dilakukan sebelum mengambil ScriptLock. Jika data unit
    // berubah sebelum commit, optimistic revision check akan menolak transaksi
    // dan file yang belum terpakai dibuang pada finally.
    photoFile = saveOdometerPhoto_(preparedPhoto, preflight.nopol);

    const lock = acquireScriptLock_(25000, 'pembaruan unit');
    try {
      // Periode dapat berakhir ketika foto sedang diunggah atau request menunggu
      // lock. Verifikasi ulang tepat sebelum transaksi ditulis ke Sheet.
      assertCalibrationSubmissionOpen_();
      const match = getMasterUnitUpdateTarget_(sheet, payload.nopol);
      const indexes = match.indexes;
      const operatingPoint = String(match.row[indexes.operatingPoint] || '');
      if (!canAccessOperatingPoint_(actor, operatingPoint)) {
        throw new Error('Unit berada di luar scope Operating Point akun Anda.');
      }
      const unit = { snapshot: unitKmSnapshot_(match.row, indexes) };
      let data;
      try {
        data = validateUnitKmUpdate_(payload, unit, dateOnlyKey_(new Date()), preparedPhoto);
      } catch (error) {
        return { success: false, errors: [{ message: error.message }] };
      }

      const log = ensureUnitKmHistorySheet_(ss);
      const statusLog = getOrCreateSheet_(ss, UNIT_STATUS_LOG_SHEET);
      const status3LLog = getOrCreateSheet_(ss, UNIT_STATUS_3L_LOG_SHEET);
      const logMap = ensureHistoryHeaders_(log, UNIT_KM_LOG_HEADERS);
      const statusLogMap = ensureHistoryHeaders_(statusLog, UNIT_STATUS_LOG_HEADERS);
      const status3LLogMap = ensureHistoryHeaders_(status3LLog, UNIT_STATUS_3L_LOG_HEADERS);
      const previousStatus = String(match.row[indexes.statusUnit] || '');
      const previousStatus3L = String(match.row[indexes.status3L] || '');
      const statusChanged = normalizeLookupText_(previousStatus) !== normalizeLookupText_(data.statusUnit);
      ensureSheetCapacity_(log, 1);
      if (statusChanged) ensureSheetCapacity_(statusLog, 1);
      ensureSheetCapacity_(status3LLog, 1);
      const nopol = match.nopol;
      const now = new Date();
      const updateId = Utilities.getUuid();
      const openingHistoryId = newCalibrationHistoryId_(now);
      const calibrationNumber = nextCalibrationNumber_(
        log, logMap, nopol, unit.snapshot.calibrationNumber
      );
      Object.assign(data, { kmPhotoId: photoFile.getId(), kmPhotoAddedAt: now,
        kmPhotoAddedBy: safeSheetText_(actor.email), kmUpdatedAt: now,
        kmUpdatedBy: safeSheetText_(actor.email), kmDate: parseDateInput_(data.kmDate),
        calibrationNumber: calibrationNumber, status3L: waitingStatus3L,
        status3LUpdatedAt: now, status3LUpdatedBy: 'SYSTEM', status3LNote: '' });
      const fieldValues = {};
      Object.keys(UNIT_KM_FIELDS)
        .concat(['calibrationNumber', 'status3L', 'status3LUpdatedAt', 'status3LUpdatedBy', 'status3LNote']).forEach(field => {
        fieldValues[field] = data[field];
      });
      const writes = buildUnitFieldWrites_(sheet, match.rowNumber, indexes, fieldValues);
      const logRow = log.getLastRow() + 1;
      const statusRow = statusLog.getLastRow() + 1;
      const status3LRow = status3LLog.getLastRow() + 1;
      const applied = [];
      try {
        writes.forEach(write => { applied.push(write); write.range.setValues(write.next); });
        log.getRange(logRow, 1, 1, log.getLastColumn()).setValues([buildHistoryRow_(log, logMap, {
          'UPDATE ID': updateId,
          'HISTORY ID': openingHistoryId,
          'REQUEST ID': updateId,
          'CALIBRATION SEQUENCE': calibrationNumber,
          'RECORD TYPE': 'CALIBRATION_OPENED',
          'LIFECYCLE STATUS': 'ACTIVE',
          'TIMESTAMP': now,
          'USER EMAIL': safeSheetText_(actor.email),
          'NOMOR POLISI': safeSheetText_(nopol),
          'KALIBRASI KE': calibrationNumber,
          'TANGGAL KM': data.kmDate,
          'STATUS UNIT': data.statusUnit,
          'KM ODOMETER': data.kmOdometer,
          'CATATAN': safeSheetText_(data.note),
          'DATA SEBELUM': JSON.stringify(unit.snapshot),
          'FOTO ODOMETER ID': data.kmPhotoId,
          'FOTO DITAMBAHKAN AT': now,
          'FOTO DIUNGGAH OLEH': safeSheetText_(actor.email),
          'OPERATING POINT': safeSheetText_(operatingPoint),
          'NAMA TOH/WOH': safeSheetText_(match.row[indexes.toh]),
          'STATUS 3L AWAL': waitingStatus3L,
          'STATUS 3L UPDATED AT': now,
          'STATUS 3L UPDATED BY': 'SYSTEM',
          'SOURCE': 'DASHBOARD',
          'REASON': 'CALIBRATION_CREATED'
        })]);
        if (statusChanged) statusLog.getRange(statusRow, 1, 1, statusLog.getLastColumn()).setValues([buildHistoryRow_(statusLog, statusLogMap, {
          'UPDATE ID': updateId,
          'TIMESTAMP': now,
          'USER EMAIL': safeSheetText_(actor.email),
          'NOMOR POLISI': safeSheetText_(nopol),
          'KALIBRASI KE': calibrationNumber,
          'STATUS SEBELUM': safeSheetText_(previousStatus),
          'STATUS BARU': data.statusUnit,
          'CATATAN': safeSheetText_(data.note),
          'SOURCE': 'DASHBOARD',
          'REASON': 'CALIBRATION_CREATED'
        })]);
        status3LLog.getRange(status3LRow, 1, 1, status3LLog.getLastColumn()).setValues([buildHistoryRow_(status3LLog, status3LLogMap, {
          'UPDATE ID': updateId,
          'TIMESTAMP': now,
          'USER EMAIL': 'SYSTEM',
          'NOMOR POLISI': safeSheetText_(nopol),
          'KALIBRASI KE': calibrationNumber,
          'STATUS SEBELUM': safeSheetText_(previousStatus3L),
          'STATUS BARU': waitingStatus3L,
          'CATATAN': safeSheetText_('Kalibrasi ke-' + calibrationNumber + ' dibuat dari pembaruan odometer dan foto.'),
          'SOURCE': 'AUTO_TRIGGER',
          'REASON': 'CALIBRATION_CREATED',
          'STATUS UNIT SEBELUM': safeSheetText_(previousStatus),
          'STATUS UNIT SESUDAH': data.statusUnit
        })]);
        SpreadsheetApp.flush();
        committed = true;
        invalidateAppCaches_();
      } catch (error) {
        try {
          applied.reverse().forEach(write => write.range.setValues(write.old));
          log.getRange(logRow, 1, 1, log.getLastColumn()).clearContent();
          if (statusChanged) statusLog.getRange(statusRow, 1, 1, statusLog.getLastColumn()).clearContent();
          status3LLog.getRange(status3LRow, 1, 1, status3LLog.getLastColumn()).clearContent();
          SpreadsheetApp.flush();
        } catch (rollbackError) {
          committed = true; // Pertahankan bukti apabila pemulihan hanya selesai sebagian.
          throw new Error('Pemulihan belum lengkap. Periksa MASTER_UNIT dan histori. Update ID: ' + updateId);
        }
        throw new Error('Update gagal; nilai sebelumnya dipulihkan. ' + error.message);
      }
      runNonCritical_('sync clean analytics row pembukaan kalibrasi', () => upsertAnalyticsCalibrationRows_([[
        updateId, now, safeSheetText_(actor.email), safeSheetText_(operatingPoint), safeSheetText_(match.row[indexes.toh]),
        safeSheetText_(nopol), safeSheetText_(data.statusUnit), calibrationNumber, data.kmDate, data.kmOdometer,
        data.kmPhotoId, now, safeSheetText_(actor.email), safeSheetText_(waitingStatus3L), safeSheetText_(waitingStatus3L),
        'SYSTEM', now, '', '', safeSheetText_(openingHistoryId), safeSheetText_(updateId), calibrationNumber
      ]]));
      runNonCritical_('audit update unit', () => logSystemActivity_('', nopol, 'UNIT_UPDATED',
        'STATUS / ODOMETER / FOTO', JSON.stringify(unit.snapshot),
        JSON.stringify({ statusUnit: data.statusUnit, kmOdometer: data.kmOdometer,
          kmPhotoId: data.kmPhotoId, calibrationNumber: calibrationNumber,
          status3L: waitingStatus3L, status3LUpdatedAt: now,
          status3LUpdatedBy: 'SYSTEM', status3LNote: '' }),
        data.note, actor, true));
      const updatedRow = sheet.getRange(match.rowNumber, 1, 1, match.readColumnCount).getValues()[0];
      return { success: true, updated: 1, calibrationNumber: calibrationNumber,
        status3L: waitingStatus3L, status3LReset: true,
        unit: buildMasterUnitClientPatch_(updatedRow, indexes) };
    } finally {
      lock.releaseLock();
    }
  } finally {
    if (!committed && photoFile) runNonCritical_('hapus foto update gagal', () => photoFile.setTrashed(true));
  }
}


function isFinalStatus3L_(status3L) {
  return normalizeLookupText_(status3L) === 'DONE';
}

/**
 * Automatic operational-status changes must never reopen a completed
 * calibration. DONE is sticky for the active calibration and may only be
 * replaced when a NEW odometer/photo update creates the next calibration.
 */
function nextAutomaticStatus3LForUnit_(currentStatus3L, nextStatusUnit, calibrationNumber) {
  const current = displayValue_(currentStatus3L);
  if (isFinalStatus3L_(current)) {
    return {
      status3L: current,
      changed: false,
      reason: 'CALIBRATION_DONE_PRESERVED',
      note: 'Status 3L DONE dipertahankan; perubahan Status Unit tidak membuka ulang kalibrasi yang sudah selesai.'
    };
  }
  const unitKey = normalizeLookupText_(nextStatusUnit);
  if (unitKey === normalizeLookupText_('On Delivery')) {
    return {
      status3L: '',
      changed: Boolean(cleanText_(current)),
      reason: 'UNIT_ON_DELIVERY',
      note: 'Status 3L dikosongkan otomatis karena unit On Delivery.'
    };
  }
  if (unitKey === normalizeLookupText_('Standby')) {
    // Jangan menimpa PROCESS/KENDALA/status aktif. WAITING hanya dibuat bila
    // status memang kosong pada siklus yang sudah memiliki nomor kalibrasi.
    if (!cleanText_(current) && parseCalibrationNumber_(calibrationNumber) > 0) {
      const waiting = resolveStatus3L_('WAITING');
      return {
        status3L: waiting,
        changed: true,
        reason: 'UNIT_STANDBY',
        note: 'Status 3L menjadi WAITING karena unit Standby dan siklus kalibrasi aktif belum selesai.'
      };
    }
  }
  return { status3L: current, changed: false, reason: '', note: '' };
}

function updateMasterUnitStatus_(payload, sessionToken) {
  const actor = requirePermission_('UPDATE_UNIT', '', { token: sessionToken });
  if (!payload || typeof payload !== 'object') {
    throw new Error('Format update status tidak valid.');
  }

  const lock = acquireScriptLock_(25000, 'pembaruan status unit');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);

    if (!sheet || sheet.getLastRow() < 2) {
      throw new Error('MASTER_UNIT belum tersedia.');
    }

    const match = getMasterUnitUpdateTarget_(sheet, payload.nopol);
    const indexes = match.indexes;
    const unit = {
      row: match.rowNumber,
      nopol: match.nopol,
      operatingPoint: displayValue_(match.row[indexes.operatingPoint]),
      raw: match.row,
      snapshot: unitKmSnapshot_(match.row, indexes)
    };
    requirePermission_('UPDATE_UNIT', unit.operatingPoint);
    if (typeof payload.kmRevision !== 'string' || payload.kmRevision !== unit.snapshot.kmRevision) {
      throw new Error('Data unit sudah berubah. Refresh Pengajuan Kalibrasi lalu ulangi update.');
    }

    const status = UNIT_MANUAL_OPERATIONAL_STATUSES.find(item =>
      normalizeLookupText_(item) === normalizeLookupText_(payload.statusUnit)
    );

    if (!status) {
      throw new Error('Status manual hanya dapat diubah ke Standby atau On Delivery. Ready to Depart dibuat otomatis saat Status 3L DONE.');
    }

    const note = String(payload.note || '').trim();
    if (note.length > 500) throw new Error('Catatan status maksimal 500 karakter.');

    const previousStatus = indexes.statusUnit >= 0
      ? displayValue_(unit.raw[indexes.statusUnit])
      : '';

    const statusChanged = normalizeLookupText_(previousStatus) !== normalizeLookupText_(status);
    const previousStatus3L = displayValue_(unit.raw[indexes.status3L]);
    const previousStatus3LNote = displayValue_(unit.raw[indexes.status3LNote]);
    const calibrationCompleted = isFinalStatus3L_(previousStatus3L) ||
      isCalibrationCompleted_(unit.nopol, unit.snapshot.calibrationNumber);
    const transition3L = calibrationCompleted
      ? { status3L: previousStatus3L, changed: false, reason: 'CALIBRATION_DONE_PRESERVED',
          note: 'Siklus kalibrasi sudah selesai; Status 3L tidak diubah oleh perubahan Status Unit.' }
      : nextAutomaticStatus3LForUnit_(previousStatus3L, status, unit.snapshot.calibrationNumber);
    const status3L = transition3L.status3L;
    const status3LChanged = Boolean(transition3L.changed);
    if (!statusChanged && !status3LChanged) {
      return { success: true, updated: 0, unchanged: true, statusUnit: status, unit: buildMasterUnitClientPatch_(unit.raw, indexes) };
    }
    const log = getOrCreateSheet_(ss, UNIT_STATUS_LOG_SHEET);
    const status3LLog = getOrCreateSheet_(ss, UNIT_STATUS_3L_LOG_SHEET);
    ensureHistoryHeaders_(log, UNIT_STATUS_LOG_HEADERS);
    ensureHistoryHeaders_(status3LLog, UNIT_STATUS_3L_LOG_HEADERS);
    if (statusChanged) ensureSheetCapacity_(log, 1);
    if (status3LChanged) ensureSheetCapacity_(status3LLog, 1);
    const logStartRow = log.getLastRow() + 1;
    const status3LLogStartRow = status3LLog.getLastRow() + 1;
    const now = new Date();
    const updateId = Utilities.getUuid();
    const changedFields = {};
    if (statusChanged) changedFields.statusUnit = status;
    if (status3LChanged) Object.assign(changedFields, {
      status3L: status3L, status3LUpdatedAt: now, status3LUpdatedBy: 'SYSTEM', status3LNote: ''
    });
    const writes = buildUnitFieldWrites_(sheet, unit.row, indexes, changedFields);
    const applied = [];

    try {
      writes.forEach(write => { applied.push(write); write.range.setValues(write.next); });
      if (statusChanged) appendStatusUnitHistoryRows_([{
        updateId: updateId, timestamp: now, email: actor.email, nopol: unit.nopol,
        calibrationNumber: unit.snapshot.calibrationNumber,
        previousStatus: previousStatus, statusUnit: status, note: note,
        source: 'DASHBOARD', reason: 'MANUAL_STATUS_UPDATE'
      }]);
      if (status3LChanged) appendStatus3LHistoryRows_([{
        updateId: updateId, timestamp: now, email: 'SYSTEM', nopol: unit.nopol,
        calibrationNumber: unit.snapshot.calibrationNumber,
        previousStatus: previousStatus3L, status3L: status3L,
        note: transition3L.note,
        source: 'AUTO_TRIGGER', reason: transition3L.reason,
        previousStatusUnit: previousStatus, statusUnit: status
      }]);
      SpreadsheetApp.flush();
      invalidateAppCaches_();
    } catch (error) {
      try {
        applied.reverse().forEach(write => write.range.setValues(write.old));
        clearAppendedHistoryRows_(log, logStartRow, statusChanged ? 1 : 0);
        clearAppendedHistoryRows_(status3LLog, status3LLogStartRow, status3LChanged ? 1 : 0);
        SpreadsheetApp.flush();
      } catch (rollbackError) {
        throw new Error('Update status terhenti dan pemulihan belum lengkap. Periksa MASTER_UNIT beserta sheet history.');
      }
      throw new Error('Update status gagal; nilai sebelumnya dipulihkan. ' + error.message);
    }

    runNonCritical_('activity log status unit', () => logSystemActivity_('', unit.nopol,
      statusChanged ? 'UNIT_STATUS_UPDATED' : 'UNIT_STATUS_3L_REPAIRED', 'STATUS UNIT / STATUS 3L',
      JSON.stringify({ statusUnit: previousStatus, status3L: previousStatus3L }),
      JSON.stringify({ statusUnit: status, status3L: status3L }), note, actor, true));

    const updatedRow = sheet.getRange(unit.row, 1, 1, match.readColumnCount).getValues()[0];
    runNonCritical_('sync analytics setelah update Status Unit', () => {
      const state = analyticsStateFromMasterRow_(updatedRow, indexes);
      if (state) updateAnalyticsCalibrationState_([state]);
    });
    return { success: true, updated: 1, statusUnit: status, status3L: status3L,
      status3LReset: status3LChanged, status3LCleared: status3LChanged && !status3L,
      unit: buildMasterUnitClientPatch_(updatedRow, indexes) };
  } finally {
    lock.releaseLock();
  }
}


function clearAppendedHistoryRows_(sheet, startRow, expectedRows) {
  if (!sheet || !expectedRows) return;
  const written = Math.min(expectedRows, Math.max(0, sheet.getLastRow() - startRow + 1));
  if (written > 0) sheet.getRange(startRow, 1, written, sheet.getLastColumn()).clearContent();
}


function appendStatusUnitHistoryRows_(entries) {
  const rows = Array.isArray(entries) ? entries.filter(Boolean) : [];
  if (!rows.length) return;
  const sheet = getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), UNIT_STATUS_LOG_SHEET);
  const map = ensureHistoryHeaders_(sheet, UNIT_STATUS_LOG_HEADERS);
  ensureSheetCapacity_(sheet, rows.length);
  const values = rows.map(entry => buildHistoryRow_(sheet, map, {
    'UPDATE ID': entry.updateId || Utilities.getUuid(),
    'TIMESTAMP': entry.timestamp || new Date(),
    'USER EMAIL': safeSheetText_(entry.email),
    'NOMOR POLISI': safeSheetText_(entry.nopol),
    'KALIBRASI KE': parseCalibrationNumber_(entry.calibrationNumber) || '',
    'STATUS SEBELUM': safeSheetText_(entry.previousStatus),
    'STATUS BARU': safeSheetText_(entry.statusUnit),
    'CATATAN': safeSheetText_(entry.note),
    'SOURCE': safeSheetText_(entry.source || 'DASHBOARD'),
    'REASON': safeSheetText_(entry.reason || 'STATUS_UNIT_UPDATED')
  }));
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, sheet.getLastColumn()).setValues(values);
}


function appendStatus3LHistoryRows_(entries) {
  const rows = Array.isArray(entries) ? entries.filter(Boolean) : [];
  if (!rows.length) return;
  const sheet = getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), UNIT_STATUS_3L_LOG_SHEET);
  const map = ensureHistoryHeaders_(sheet, UNIT_STATUS_3L_LOG_HEADERS);
  ensureSheetCapacity_(sheet, rows.length);
  const values = rows.map(entry => buildHistoryRow_(sheet, map, {
    'UPDATE ID': entry.updateId || Utilities.getUuid(),
    'TIMESTAMP': entry.timestamp || new Date(),
    'USER EMAIL': safeSheetText_(entry.email),
    'NOMOR POLISI': safeSheetText_(entry.nopol),
    'KALIBRASI KE': parseCalibrationNumber_(entry.calibrationNumber) || '',
    'STATUS SEBELUM': safeSheetText_(entry.previousStatus),
    'STATUS BARU': safeSheetText_(entry.status3L),
    'CATATAN': safeSheetText_(entry.note),
    'SOURCE': safeSheetText_(entry.source || 'DASHBOARD'),
    'STATUS UNIT SEBELUM': safeSheetText_(entry.previousStatusUnit),
    'STATUS UNIT SESUDAH': safeSheetText_(entry.statusUnit),
    'REASON': safeSheetText_(entry.reason || 'STATUS_3L_UPDATED')
  }));
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, sheet.getLastColumn()).setValues(values);
}


/** Menulis snapshot final tanpa menimpa row pembukaan siklus. */
function appendCompletedCalibrationSnapshots_(entries) {
  const completed = (Array.isArray(entries) ? entries : []).filter(entry =>
    entry && normalizeLookupText_(entry.status3L) === 'DONE'
  );
  if (!completed.length) return null;

  const sheet = ensureUnitKmHistorySheet_(SpreadsheetApp.getActiveSpreadsheet());
  const map = ensureHistoryHeaders_(sheet, UNIT_KM_LOG_HEADERS);
  if (sheet.getLastRow() < 2) {
    throw new Error('CALIBRATION_HISTORY_MISSING: Riwayat pembukaan kalibrasi belum tersedia.');
  }
  const existing = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const baseByKey = Object.create(null);
  const finalKeys = new Set();
  existing.forEach(row => {
    const type = map['RECORD TYPE'] ? normalizeLookupText_(row[map['RECORD TYPE'] - 1]) : '';
    const key = normalizeNopol_(row[map['NOMOR POLISI'] - 1]) + '|' +
      parseCalibrationNumber_(row[map['KALIBRASI KE'] - 1]);
    if (key === '|0') return;
    const lifecycle = normalizeLookupText_(getRowValue_(row, map, 'LIFECYCLE STATUS'));
    if (type === 'FINAL_SNAPSHOT' && lifecycle === 'LOCKED') {
      finalKeys.add(key);
      return;
    }
    baseByKey[key] = row;
  });

  const pending = completed.filter(entry => {
    const key = calibrationCompletionKey_(entry.nopol, entry.calibrationNumber);
    if (!key || finalKeys.has(key)) return false;
    finalKeys.add(key); // Deduplicate juga entries pada batch yang sama.
    return true;
  });
  if (!pending.length) return { sheet: sheet, startRow: 0, count: 0, skippedExisting: completed.length };

  const rows = pending.map(entry => {
    const sequence = parseCalibrationNumber_(entry.calibrationNumber);
    const key = normalizeNopol_(entry.nopol) + '|' + sequence;
    const base = baseByKey[key];
    if (!base) {
      throw new Error('CALIBRATION_HISTORY_MISSING: Siklus ' + entry.nopol +
        ' kalibrasi ke-' + sequence + ' tidak ditemukan.');
    }
    const timestamp = entry.timestamp instanceof Date ? entry.timestamp : new Date();
    const row = base.slice();
    const businessActor = normalizeDoneStatus3LActor_('DONE', entry.email);
    const finalState = Object.assign({}, entry.dataAfter || {}, {
      statusUnit: 'Ready to Depart',
      status3L: 'DONE',
      status3LUpdatedBy: businessActor
    });
    setRowValue_(row, map, 'UPDATE ID', entry.updateId || Utilities.getUuid());
    setRowValue_(row, map, 'HISTORY ID', newCalibrationHistoryId_(timestamp));
    setRowValue_(row, map, 'REQUEST ID', cleanText_(getRowValue_(base, map, 'REQUEST ID')) ||
      cleanText_(getRowValue_(base, map, 'UPDATE ID')));
    setRowValue_(row, map, 'CALIBRATION SEQUENCE', sequence);
    setRowValue_(row, map, 'RECORD TYPE', 'FINAL_SNAPSHOT');
    setRowValue_(row, map, 'LIFECYCLE STATUS', 'LOCKED');
    setRowValue_(row, map, 'TIMESTAMP', timestamp);
    setRowValue_(row, map, 'USER EMAIL', safeSheetText_(entry.email));
    setRowValue_(row, map, 'STATUS UNIT', safeSheetText_(finalState.statusUnit || 'Ready to Depart'));
    // Kolom scalar dan DATA SETELAH harus menunjuk bukti yang sama.
    if (normalizeOdometerPhotoId_(finalState.kmPhotoId)) {
      setRowValue_(row, map, 'FOTO ODOMETER ID', normalizeOdometerPhotoId_(finalState.kmPhotoId));
      if (finalState.kmPhotoAddedAt) setRowValue_(row, map, 'FOTO DITAMBAHKAN AT', parseDateTimeInput_(finalState.kmPhotoAddedAt) || finalState.kmPhotoAddedAt);
      if (finalState.kmPhotoAddedBy) setRowValue_(row, map, 'FOTO DIUNGGAH OLEH', safeSheetText_(finalState.kmPhotoAddedBy));
    }
    if (finalState.kmOdometer !== undefined) setRowValue_(row, map, 'KM ODOMETER', finalState.kmOdometer);
    if (finalState.kmDate) setRowValue_(row, map, 'TANGGAL KM', finalState.kmDate);
    setRowValue_(row, map, 'STATUS 3L UPDATED AT', timestamp);
    setRowValue_(row, map, 'STATUS 3L UPDATED BY', safeSheetText_(businessActor));
    setRowValue_(row, map, 'COMPLETED AT', timestamp);
    setRowValue_(row, map, 'COMPLETED BY', safeSheetText_(businessActor));
    setRowValue_(row, map, 'SOURCE', safeSheetText_(entry.source || 'DASHBOARD'));
    setRowValue_(row, map, 'REASON', safeSheetText_(entry.reason || 'STATUS_3L_DONE'));
    setRowValue_(row, map, 'DATA SETELAH', JSON.stringify(finalState));
    return row;
  });
  ensureSheetCapacity_(sheet, rows.length);
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, sheet.getLastColumn()).setValues(rows);
  return { sheet: sheet, startRow: startRow, count: rows.length };
}



/**
 * Melengkapi event Status 3L yang hilang pada database legacy tanpa mengubah
 * event lama. Hanya menambahkan event rekonsiliasi bila state aktif MASTER_UNIT
 * berbeda dengan event terakhir untuk nopol + nomor kalibrasi yang sama.
 */

/**
 * Backfill Status Unit hanya dilakukan bila tersedia evidence kuat dari
 * workflow DONE => Ready to Depart. Untuk status manual lain, Summary tetap
 * dapat melakukan current-state reconciliation tanpa menciptakan timestamp audit semu.
 */
function backfillCurrentStatusUnitHistoryFromMaster_(masterSheet, indexes) {
  // Deprecated v21: intentionally disabled. MASTER_UNIT must never create the
  // history evidence later used to validate MASTER_UNIT itself.
  return { written: 0, skippedUnsafeMasterMirroring: true };
}

function backfillCurrentStatus3LHistoryFromMaster_(masterSheet, indexes, status3LSheet) {
  // Deprecated v21: intentionally disabled to prevent circular self-validation.
  return { written: 0, skippedUnsafeMasterMirroring: true };
}

function backfillCompletedCalibrationSnapshots_(status3LSheet) {
  const history = getUnitKmHistorySheet_(SpreadsheetApp.getActiveSpreadsheet());
  if (!history || history.getLastRow() < 2 || !status3LSheet || status3LSheet.getLastRow() < 2) return 0;
  const trustedStatus = buildTrustedStatus3LStateIndex_(status3LSheet);
  const trustedFinals = buildTrustedFinalSnapshotIndex_(trustedStatus);
  const historyMap = ensureHistoryHeaders_(history, UNIT_KM_LOG_HEADERS);
  const baseKeys = new Set();
  history.getRange(2, 1, history.getLastRow() - 1, history.getLastColumn()).getValues().forEach(row => {
    const type = normalizeLookupText_(row[historyMap['RECORD TYPE'] - 1]);
    const lifecycle = normalizeLookupText_(row[historyMap['LIFECYCLE STATUS'] - 1]);
    const key = calibrationCompletionKey_(row[historyMap['NOMOR POLISI'] - 1], row[historyMap['KALIBRASI KE'] - 1]);
    if (!key) return;
    if (type !== 'FINAL_SNAPSHOT' || lifecycle !== 'LOCKED') baseKeys.add(key);
  });
  const entries = Object.keys(trustedStatus).map(key => {
    const bucket = trustedStatus[key];
    const done = bucket && bucket.done;
    if (!done || trustedFinals[key] || !baseKeys.has(key)) return null;
    const parts = key.split('|');
    const sequence = Number(parts.pop()) || 0;
    const nopol = parts.join('|');
    return {
      updateId: Utilities.getUuid(),
      timestamp: done.timestamp || new Date(),
      email: done.updatedBy || KM_STATUS_3L_SHEET_FALLBACK,
      nopol: nopol,
      calibrationNumber: sequence,
      status3L: 'DONE',
      source: 'SYSTEM_REPAIR',
      reason: 'REPAIR_FINAL_SNAPSHOT_FROM_TRUSTED_DONE',
      dataAfter: {
        statusUnit: 'Ready to Depart',
        status3L: 'DONE',
        status3LNote: done.note || '',
        status3LUpdatedAt: done.timestamp ? done.timestamp.toISOString() : '',
        status3LUpdatedBy: done.updatedBy || KM_STATUS_3L_SHEET_FALLBACK,
        calibrationNumber: sequence
      }
    };
  }).filter(Boolean);
  const written = appendCompletedCalibrationSnapshots_(entries);
  return written ? written.count : 0;
}

function buildLatestStatusUnitMap_() {
  const result = Object.create(null);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNIT_STATUS_LOG_SHEET);
  if (!sheet || sheet.getLastRow() < 2 || !historyHeadersReady_(sheet, UNIT_STATUS_LOG_HEADERS)) return result;
  const map = getHistoryHeaderMap_(sheet, UNIT_STATUS_LOG_HEADERS);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  for (let index = rows.length - 1; index >= 0; index--) {
    const key = normalizeNopol_(rows[index][map['NOMOR POLISI'] - 1]);
    if (key && result[key] === undefined) result[key] = displayValue_(rows[index][map['STATUS BARU'] - 1]);
  }
  return result;
}


function buildLatestStatus3LMap_() {
  const result = Object.create(null);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNIT_STATUS_3L_LOG_SHEET);
  const trusted = buildTrustedStatus3LStateIndex_(sheet);
  Object.keys(trusted).forEach(key => {
    const latest = trusted[key] && trusted[key].latest;
    if (latest) result[key] = displayValue_(latest.status3L);
  });
  return result;
}


function updateMasterUnitStatus3LBulk(payload, sessionToken) {
  const actor = requirePermission_('UPDATE_STATUS_3L', '', { token: sessionToken });
  if (!payload || typeof payload !== 'object' || !Array.isArray(payload.nopols)) {
    throw new Error('Format bulk update Status 3L tidak valid.');
  }
  const requested = [...new Set(payload.nopols.map(normalizeDisplayNopol_).filter(Boolean))];
  if (!requested.length) throw new Error('Pilih minimal satu unit.');
  if (requested.length > STATUS_3L_WEB_BULK_MAX_UNITS) {
    throw new Error('Maksimal ' + STATUS_3L_WEB_BULK_MAX_UNITS +
      ' unit per proses bulk melalui dashboard. Untuk volume lebih besar, lakukan langsung pada MASTER_UNIT.');
  }
  const status3L = resolveStatus3L_(payload.status3L);
  const note = String(payload.note || '').trim();
  if (note.length > 500) throw new Error('Catatan Status 3L maksimal 500 karakter.');

  const lock = acquireScriptLock_(25000, 'pembaruan Status 3L');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
    if (!sheet || sheet.getLastRow() < 2) throw new Error('MASTER_UNIT belum tersedia.');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
    if (indexes.nopol < 0 || indexes.operatingPoint < 0 || indexes.status3L < 0 ||
        indexes.kmOdometer < 0 || indexes.kmDate < 0 || indexes.kmPhotoId < 0 ||
        indexes.calibrationNumber < 0 || indexes.status3LUpdatedAt < 0 ||
        indexes.status3LUpdatedBy < 0 || indexes.status3LNote < 0 ||
        indexes.status3L !== indexes.calibrationNumber + 1 ||
        indexes.status3LUpdatedAt !== indexes.status3L + 1 ||
        indexes.status3LUpdatedBy !== indexes.status3LUpdatedAt + 1 ||
        indexes.status3LNote !== indexes.status3LUpdatedBy + 1) {
      throw new Error('Database belum menggunakan struktur terbaru. SUPER_ADMIN perlu menjalankan Perbaiki Database / Setup Migrasi sebelum bulk Status 3L digunakan.');
    }
    if (!isMasterUnitStatusValidationReady_(sheet, indexes.statusUnit)) {
      // Self-heal rule lama yang hanya berisi Standby/On Delivery. Dengan ini
      // transaksi dashboard tidak gagal hanya karena deployment sebelumnya belum
      // sempat menjalankan maintenance validation.
      applyMasterUnitStatusValidation_(sheet, indexes.statusUnit);
      if (!isMasterUnitStatusValidationReady_(sheet, indexes.statusUnit)) {
        throw new Error('Database belum siap menulis Ready to Depart otomatis. SUPER_ADMIN perlu menjalankan Repair Workflow 3L.');
      }
    }

    // Satu batch read jauh lebih cepat daripada enam round-trip Range terpisah.
    const rowCount = sheet.getLastRow() - 1;
    const masterRows = sheet.getRange(2, 1, rowCount, sheet.getLastColumn()).getValues();
    const directory = Object.create(null);
    masterRows.forEach((row, offset) => {
      const key = normalizeNopol_(row[indexes.nopol]);
      if (!key) return;
      if (directory[key]) directory[key].duplicate = true;
      else {
        const eligibility = unitStatus3LEligibility_({
          kmOdometer: row[indexes.kmOdometer],
          kmDate: row[indexes.kmDate],
          kmPhotoId: row[indexes.kmPhotoId],
          calibrationNumber: row[indexes.calibrationNumber],
          statusUnit: row[indexes.statusUnit],
          status3L: row[indexes.status3L]
        });
        directory[key] = {
          rowNumber: offset + 2,
          rawRow: row.slice(),
          nopol: normalizeDisplayNopol_(row[indexes.nopol]),
          operatingPoint: displayValue_(row[indexes.operatingPoint]),
          previousStatus: displayValue_(row[indexes.status3L]),
          previousStatus3LUpdatedAt: row[indexes.status3LUpdatedAt],
          previousStatus3LUpdatedBy: displayValue_(row[indexes.status3LUpdatedBy]),
          previousNote: displayValue_(row[indexes.status3LNote]),
          statusUnit: displayValue_(row[indexes.statusUnit]),
          calibrationNumber: parseCalibrationNumber_(row[indexes.calibrationNumber]),
          status3LEligible: eligibility.eligible,
          status3LEligibilityReason: eligibility.reason,
          duplicate: false
        };
      }
    });

    const targets = requested.map(nopol => {
      const item = directory[normalizeNopol_(nopol)];
      if (!item) throw new Error('Unit tidak ditemukan: ' + nopol);
      if (item.duplicate) throw new Error('Nomor polisi duplikat di MASTER_UNIT: ' + nopol);
      if (!canAccessOperatingPoint_(actor, item.operatingPoint)) {
        throw new Error('Akses ditolak untuk unit ' + nopol + ' (' + item.operatingPoint + ').');
      }
      return item;
    });
    const finalizingDone = normalizeLookupText_(status3L) === 'DONE';
    const businessActor = normalizeDoneStatus3LActor_(status3L, actor.email);
    const ineligible = targets.filter(item => !item.status3LEligible &&
      !(finalizingDone && item.calibrationNumber > 0));
    if (ineligible.length) {
      const examples = ineligible.slice(0, 5).map(item => item.nopol + ': ' + item.status3LEligibilityReason).join(' | ');
      throw new Error('STATUS_3L_NOT_ELIGIBLE: ' + examples +
        (ineligible.length > 5 ? ' | dan ' + (ineligible.length - 5) + ' unit lainnya.' : ''));
    }
    const changed = targets.filter(item =>
      normalizeLookupText_(item.previousStatus) !== normalizeLookupText_(status3L) ||
      cleanText_(item.previousNote) !== note ||
      (finalizingDone &&
        normalizeLookupText_(item.statusUnit) !== normalizeLookupText_('Ready to Depart')) ||
      (finalizingDone &&
        normalizeLookupText_(item.previousStatus3LUpdatedBy) !== normalizeLookupText_(KM_STATUS_3L_SHEET_FALLBACK))
    );
    if (!changed.length) {
      return { success: true, updated: 0, unchanged: requested.length, status3L, note, nopols: [], units: [] };
    }

    const statusA1 = changed.map(item => sheet.getRange(item.rowNumber, indexes.status3L + 1).getA1Notation());
    const timestampA1 = changed.map(item => sheet.getRange(item.rowNumber, indexes.status3LUpdatedAt + 1).getA1Notation());
    const updatedByA1 = changed.map(item => sheet.getRange(item.rowNumber, indexes.status3LUpdatedBy + 1).getA1Notation());
    const noteA1 = changed.map(item => sheet.getRange(item.rowNumber, indexes.status3LNote + 1).getA1Notation());
    const logSheet = getOrCreateSheet_(ss, UNIT_STATUS_3L_LOG_SHEET);
    ensureHistoryHeaders_(logSheet, UNIT_STATUS_3L_LOG_HEADERS);
    const statusLogSheet = getOrCreateSheet_(ss, UNIT_STATUS_LOG_SHEET);
    ensureHistoryHeaders_(statusLogSheet, UNIT_STATUS_LOG_HEADERS);
    changed.forEach(item => {
      // Satu arah: DONE -> Ready to Depart. Status 3L lain tidak melakukan
      // rollback operational status dan tidak mencari previous status dari log.
      item.nextStatusUnit = normalizeLookupText_(status3L) === 'DONE'
        ? 'Ready to Depart'
        : item.statusUnit;
      item.statusUnitChanged = normalizeLookupText_(item.statusUnit) !== normalizeLookupText_(item.nextStatusUnit);
      item.updateId = Utilities.getUuid();
    });
    const unitStatusChanged = changed.filter(item => item.statusUnitChanged);
    const statusUnitRangesByValue = Object.create(null);
    unitStatusChanged.forEach(item => {
      const value = item.nextStatusUnit;
      if (!statusUnitRangesByValue[value]) statusUnitRangesByValue[value] = [];
      statusUnitRangesByValue[value].push(
        sheet.getRange(item.rowNumber, indexes.statusUnit + 1).getA1Notation()
      );
    });
    ensureSheetCapacity_(logSheet, changed.length);
    ensureSheetCapacity_(statusLogSheet, unitStatusChanged.length);
    const now = new Date();

    // PRIMARY COMMIT. Setelah cell Status 3L berhasil ditulis, kegagalan audit/history
    // tidak boleh mengembalikan DONE menjadi WAITING/previous status. Ini memisahkan
    // business state dari best-effort audit persistence.
    changed.forEach(item => {
      const live = sheet.getRange(item.rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (!masterWorkflowRowUnchanged_(item.rawRow, live, indexes)) {
        throw new Error('DATA_CHANGED: Status unit ' + item.nopol + ' berubah saat diproses. Muat ulang dan coba kembali.');
      }
    });
    try {
      sheet.getRangeList(statusA1).setValue(safeSheetText_(status3L));
      sheet.getRangeList(timestampA1).setValue(now);
      sheet.getRangeList(updatedByA1).setValue(safeSheetText_(businessActor));
      sheet.getRangeList(noteA1).setValue(safeSheetText_(note));
      Object.keys(statusUnitRangesByValue).forEach(value => {
        sheet.getRangeList(statusUnitRangesByValue[value]).setValue(value);
      });
      SpreadsheetApp.flush();
    } catch (error) {
      // Tidak ada write-back previousStatus di sini. Jika Google gagal melakukan
      // primary commit, laporkan error apa adanya dan biarkan audit owner memeriksa
      // row yang mungkin parsial; jangan membuat rollback otomatis tersembunyi.
      throw new Error('STATUS_3L_WRITE_FAILED: pembaruan utama belum selesai. ' +
        String(error && error.message ? error.message : error));
    }

    const auditWarnings = [];
    const saveAudit = (label, task) => {
      try {
        task();
        return true;
      } catch (error) {
        const message = label + ': ' + String(error && error.message ? error.message : error);
        auditWarnings.push(message);
        console.error(message, error && error.stack ? error.stack : error);
        return false;
      }
    };

    saveAudit('STATUS_3L_HISTORY', () => appendStatus3LHistoryRows_(changed.map(item => ({
      updateId: item.updateId, timestamp: now, email: businessActor,
      nopol: item.nopol, previousStatus: item.previousStatus, status3L,
      calibrationNumber: item.calibrationNumber,
      note, source: 'DASHBOARD', reason: 'STATUS_3L_UPDATED',
      previousStatusUnit: item.statusUnit, statusUnit: item.nextStatusUnit
    }))));

    saveAudit('STATUS_UNIT_HISTORY', () => appendStatusUnitHistoryRows_(unitStatusChanged.map(item => ({
      updateId: item.updateId, timestamp: now, email: 'SYSTEM',
      nopol: item.nopol, calibrationNumber: item.calibrationNumber,
      previousStatus: item.statusUnit, statusUnit: item.nextStatusUnit,
      note: 'Ready to Depart otomatis karena Status 3L DONE.',
      source: 'AUTO_TRIGGER', reason: 'STATUS_3L_DONE'
    }))));

    if (normalizeLookupText_(status3L) === 'DONE') {
      saveAudit('FINAL_SNAPSHOT', () => appendCompletedCalibrationSnapshots_(changed
        .filter(item => normalizeLookupText_(item.previousStatus) !== 'DONE')
        .map(item => {
          const finalRow = masterRows[item.rowNumber - 2].slice();
          finalRow[indexes.status3L] = status3L;
          finalRow[indexes.status3LUpdatedAt] = now;
          finalRow[indexes.status3LUpdatedBy] = businessActor;
          finalRow[indexes.status3LNote] = note;
          finalRow[indexes.statusUnit] = item.nextStatusUnit;
          return {
            updateId: item.updateId, timestamp: now, email: businessActor,
            nopol: item.nopol, calibrationNumber: item.calibrationNumber,
            status3L: status3L, source: 'DASHBOARD', reason: 'STATUS_3L_DONE',
            dataAfter: unitKmSnapshot_(finalRow, indexes)
          };
        })));
    }

    if (auditWarnings.length) {
      runNonCritical_('catat warning audit Status 3L', () => logSystemActivity_('', '',
        'STATUS_3L_AUDIT_WARNING', 'STATUS 3L', '', status3L,
        auditWarnings.join(' | '), actor, true));
    }


    invalidateAppCaches_();
    runNonCritical_('audit bulk Status 3L', () => logSystemActivity_('', '', 'STATUS_3L_BULK_UPDATED',
      'STATUS 3L', '', status3L, changed.length + ' unit | ' + note, actor, true));
    changed.forEach(item => {
      const row = masterRows[item.rowNumber - 2];
      row[indexes.status3L] = status3L;
      row[indexes.status3LUpdatedAt] = now;
      row[indexes.status3LUpdatedBy] = businessActor;
      row[indexes.status3LNote] = note;
      row[indexes.statusUnit] = item.nextStatusUnit;
    });
    runNonCritical_('sync analytics setelah bulk Status 3L', () => updateAnalyticsCalibrationState_(changed
      .map(item => analyticsStateFromMasterRow_(masterRows[item.rowNumber - 2], indexes)).filter(Boolean)));
    return {
      success: true,
      updated: changed.length,
      unchanged: targets.length - changed.length,
      status3L,
      note,
      auditWarnings: auditWarnings,
      nopols: changed.map(item => item.nopol),
      units: changed.map(item => buildMasterUnitClientPatch_(masterRows[item.rowNumber - 2], indexes))
    };
  } finally {
    lock.releaseLock();
  }
}



function getMasterUnitKmHistory(nopol, sessionToken) {
  const actor = requirePermission_('VIEW_MASTER_UNIT', '', { token: sessionToken });
  assertUnitScope_(actor, nopol);
  return getMasterUnitKmHistory_(nopol);
}


/**
 * Membaca snapshot MASTER_UNIT saat ini untuk Ringkasan Pengajuan Kalibrasi.
 * MASTER_UNIT adalah current projection, bukan evidence completion.
 * Endpoint summary selalu memprioritaskan trusted STATUS_3L_HISTORY/final snapshot;
 * snapshot MASTER_UNIT hanya menjadi fallback display bila ledger tepercaya memang
 * belum memiliki event untuk siklus aktif.
 */
function readMasterUnitSummaryTarget_(nopol) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return null;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  if (indexes.nopol < 0) return null;
  const rowNumber = findMasterUnitRowNumberByNopol_(sheet, indexes, nopol);
  if (!rowNumber) return null;
  const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  return { headers: headers, indexes: indexes, row: row, rowNumber: rowNumber };
}

function getMasterUnitCurrentSummaryState_(nopol, target) {
  const record = target === undefined ? readMasterUnitSummaryTarget_(nopol) : target;
  if (!record) return null;
  const indexes = record.indexes;
  const row = record.row;
  const snapshot = unitKmSnapshot_(row, indexes);
  return {
    rowNumber: record.rowNumber,
    nopol: normalizeDisplayNopol_(row[indexes.nopol]),
    calibrationNumber: parseCalibrationNumber_(snapshot.calibrationNumber),
    statusUnit: displayValue_(row[indexes.statusUnit]),
    status3L: displayValue_(row[indexes.status3L]),
    status3LUpdatedAt: snapshot.status3LUpdatedAt || '',
    status3LUpdatedBy: snapshot.status3LUpdatedBy || '',
    status3LNote: snapshot.status3LNote || '',
    kmUpdatedAt: snapshot.kmUpdatedAt || '',
    kmUpdatedBy: snapshot.kmUpdatedBy || '',
    kmDate: snapshot.kmDate || '',
    kmOdometer: snapshot.kmOdometer,
    kmPhotoId: snapshot.kmPhotoId || ''
  };
}

function reconcileSummaryStatus3LHistory_(history, currentState) {
  const rows = Array.isArray(history) ? history.slice() : [];
  if (!currentState || !parseCalibrationNumber_(currentState.calibrationNumber)) {
    return { rows: rows, reconciled: false };
  }
  const sequence = parseCalibrationNumber_(currentState.calibrationNumber);
  const latestSameCycle = rows.find(item =>
    parseCalibrationNumber_(item.calibrationNumber) === sequence && cleanText_(item.status3L)
  ) || rows.find(item => parseCalibrationNumber_(item.calibrationNumber) === sequence) || null;
  const currentKey = normalizeLookupText_(currentState.status3L);
  const currentUnitKey = normalizeLookupText_(currentState.statusUnit);
  const shouldRepresentBlank = !currentKey &&
    currentUnitKey === normalizeLookupText_('On Delivery');

  if (!currentKey && !shouldRepresentBlank) return { rows: rows, reconciled: false };

  const latestKey = latestSameCycle ? normalizeLookupText_(latestSameCycle.status3L) : '';
  const sameBusinessState = latestSameCycle && latestKey === currentKey;
  if (sameBusinessState) return { rows: rows, reconciled: false };

  // MASTER_UNIT adalah primary projection untuk siklus aktif. Jika automation/API
  // mengubah STATUS 3L tanpa memicu onEdit, ledger dapat tertinggal. Summary tidak
  // boleh menampilkan WAITING ketika primary state sudah PROCESS/DONE/KENDALA.
  rows.unshift({
    updatedAt: currentState.status3LUpdatedAt || currentState.kmUpdatedAt || '',
    updatedBy: status3LActorDisplay_(currentState.status3LUpdatedBy) || KM_STATUS_3L_SHEET_FALLBACK,
    calibrationNumber: sequence,
    previousStatus: latestSameCycle ? String(latestSameCycle.status3L || '') : '',
    status3L: String(currentState.status3L || ''),
    note: currentState.status3LNote ||
      'Status aktual dari MASTER_UNIT; ledger Status 3L sedang menunggu sinkronisasi.',
    source: 'MASTER_UNIT',
    reason: 'CURRENT_STATE_RECONCILIATION',
    previousStatusUnit: latestSameCycle ? String(latestSameCycle.statusUnit || '') : '',
    statusUnit: String(currentState.statusUnit || ''),
    synthetic: true,
    currentState: true,
    trusted: false
  });
  return { rows: rows.slice(0, 100), reconciled: true };
}

function reconcileSummaryStatusUnitHistory_(history, currentState) {
  const rows = Array.isArray(history) ? history.slice() : [];
  if (!currentState || !cleanText_(currentState.statusUnit)) {
    return { rows: rows, reconciled: false };
  }
  const latest = rows.length ? rows[0] : null;
  if (latest && normalizeLookupText_(latest.statusUnit) === normalizeLookupText_(currentState.statusUnit)) {
    return { rows: rows, reconciled: false };
  }
  const completedNow = normalizeLookupText_(currentState.status3L) === 'DONE';
  rows.unshift({
    updatedAt: (completedNow ? currentState.status3LUpdatedAt : '') || currentState.kmUpdatedAt || '',
    updatedBy: status3LActorDisplay_(completedNow ? currentState.status3LUpdatedBy : '') || currentState.kmUpdatedBy || 'SYSTEM',
    calibrationNumber: parseCalibrationNumber_(currentState.calibrationNumber),
    previousStatus: latest ? String(latest.statusUnit || '') : '',
    statusUnit: String(currentState.statusUnit || ''),
    note: 'Status aktual diselaraskan dari MASTER_UNIT karena histori status unit sebelumnya belum lengkap.',
    source: 'MASTER_UNIT',
    reason: 'CURRENT_STATE_RECONCILIATION',
    synthetic: true,
    currentState: true
  });
  return { rows: rows.slice(0, 30), reconciled: true };
}

function calibrationStatusTimelineForSummary_(events) {
  // Summary uses only independently trusted events. Legacy SYSTEM_REPAIR rows
  // stay visible in the raw audit ledger but cannot freeze a cycle as DONE.
  const chronological = (events || []).filter(event => event && event.trusted !== false).slice().reverse();
  const accepted = [];
  for (let index = 0; index < chronological.length; index++) {
    const event = chronological[index];
    accepted.push(event);
    if (normalizeLookupText_(event.status3L) === 'DONE') break;
  }
  return accepted;
}


function resolveCalibrationStatusSnapshot_(events) {
  const timeline = calibrationStatusTimelineForSummary_(events);
  return timeline.length ? timeline[timeline.length - 1] : null;
}

// Ringkasan unit sudah memuat maksimal 100 event Status 3L milik unit yang sama.
// Bentuk index completion hanya dari subset tersebut agar modal tidak membaca
// seluruh STATUS_3L_HISTORY dua kali lagi pada setiap pembukaan detail.
function buildTrustedStatus3LStateIndexFromSummaryRows_(nopol, rows) {
  const result = Object.create(null);
  (rows || []).forEach((item, offset) => {
    if (!item || item.trusted === false) return;
    const key = calibrationCompletionKey_(nopol, item.calibrationNumber);
    if (!key) return;
    const parsed = item.updatedAt instanceof Date ? item.updatedAt : parseDateTimeInput_(item.updatedAt);
    const timestamp = parsed instanceof Date && !isNaN(parsed.getTime()) ? parsed : null;
    const event = {
      // rows summary datang newest-first; fallback ordering dibuat chronological.
      rowNumber: (rows || []).length - offset,
      timestamp: timestamp,
      time: timestamp ? timestamp.getTime() : -offset,
      updatedBy: status3LActorDisplay_(item.updatedBy),
      previousStatus: item.previousStatus || '',
      status3L: item.status3L || '',
      note: item.note || '',
      source: item.source || '',
      reason: item.reason || '',
      statusUnitBefore: item.previousStatusUnit || '',
      statusUnitAfter: item.statusUnit || ''
    };
    if (!result[key]) result[key] = { events: [], initial: null, done: null, latest: null };
    result[key].events.push(event);
  });
  Object.keys(result).forEach(key => {
    const bucket = result[key];
    bucket.events.sort((a, b) => a.time - b.time || a.rowNumber - b.rowNumber);
    bucket.initial = bucket.events.find(event =>
      normalizeLookupText_(event.reason) === 'CALIBRATION_CREATED' && cleanText_(event.status3L)
    ) || bucket.events.find(event => cleanText_(event.status3L)) || null;
    bucket.done = bucket.events.filter(event => normalizeLookupText_(event.status3L) === 'DONE').pop() || null;
    bucket.latest = bucket.events.length ? bucket.events[bucket.events.length - 1] : null;
  });
  return result;
}


function safeParseHistoryJson_(value) {
  if (!cleanText_(value)) return {};
  try {
    const parsed = JSON.parse(String(value));
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch (error) {
    return {};
  }
}

function buildFinalCalibrationSnapshotMap_(rows, map, trustedStatusIndex) {
  const result = Object.create(null);
  const at = header => map[header] ? map[header] - 1 : -1;
  const trustedStatus = trustedStatusIndex || buildTrustedStatus3LStateIndex_(
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNIT_STATUS_3L_LOG_SHEET));
  (rows || []).forEach(row => {
    const recordType = at('RECORD TYPE') >= 0 ? normalizeLookupText_(row[at('RECORD TYPE')]) : '';
    const lifecycle = at('LIFECYCLE STATUS') >= 0 ? normalizeLookupText_(row[at('LIFECYCLE STATUS')]) : '';
    if (recordType !== 'FINAL_SNAPSHOT' || lifecycle !== 'LOCKED') return;
    const nopol = at('NOMOR POLISI') >= 0 ? row[at('NOMOR POLISI')] : '';
    const sequence = at('KALIBRASI KE') >= 0 ? parseCalibrationNumber_(row[at('KALIBRASI KE')]) : 0;
    const key = calibrationCompletionKey_(nopol, sequence);
    const trustedDone = key && trustedStatus[key] && trustedStatus[key].done;
    if (!sequence || (!trustedDone && !isDirectTrustedFinalSnapshot_(row, map)) || result[sequence]) return;
    const dataAfter = at('DATA SETELAH') >= 0 ? safeParseHistoryJson_(row[at('DATA SETELAH')]) : {};
    const completedAtRaw = at('COMPLETED AT') >= 0 ? row[at('COMPLETED AT')] : '';
    const statusUpdatedAtRaw = at('STATUS 3L UPDATED AT') >= 0 ? row[at('STATUS 3L UPDATED AT')] : '';
    const timestampRaw = at('TIMESTAMP') >= 0 ? row[at('TIMESTAMP')] : '';
    const completedAt = completedAtRaw instanceof Date ? completedAtRaw.toISOString() :
      (statusUpdatedAtRaw instanceof Date ? statusUpdatedAtRaw.toISOString() :
        (timestampRaw instanceof Date ? timestampRaw.toISOString() : String(completedAtRaw || statusUpdatedAtRaw || timestampRaw || '')));
    const completedBy = status3LActorDisplay_(
      (at('COMPLETED BY') >= 0 ? displayValue_(row[at('COMPLETED BY')]) : '') ||
      (at('STATUS 3L UPDATED BY') >= 0 ? displayValue_(row[at('STATUS 3L UPDATED BY')]) : '') ||
      (at('USER EMAIL') >= 0 ? displayValue_(row[at('USER EMAIL')]) : '')
    );
    result[sequence] = {
      status3L: 'DONE', updatedAt: completedAt, updatedBy: completedBy,
      note: cleanText_(dataAfter.status3LNote) || '',
      statusUnit: cleanText_(dataAfter.statusUnit) ||
        (at('STATUS UNIT') >= 0 ? displayValue_(row[at('STATUS UNIT')]) : 'Ready to Depart'),
      completedAt: completedAt, completedBy: completedBy,
      recordType: recordType, lifecycleStatus: lifecycle,
      source: at('SOURCE') >= 0 ? displayValue_(row[at('SOURCE')]) : '',
      reason: at('REASON') >= 0 ? displayValue_(row[at('REASON')]) : '',
      trusted: true
    };
  });
  return result;
}

function getMasterUnitKmHistory_(nopol, status3LHistory, currentState) {
  const sheet = getUnitKmHistorySheet_(SpreadsheetApp.getActiveSpreadsheet());
  if (!sheet || sheet.getLastRow() < 2) return [];
  if (!historyHeadersReady_(sheet, UNIT_KM_LOG_HEADERS)) return [];
  const map = getHistoryHeaderMap_(sheet, UNIT_KM_LOG_HEADERS);
  const at = header => map[header] - 1;
  const statusByCalibration = Object.create(null);
  const statusTimelineByCalibration = Object.create(null);
  (status3LHistory || getMasterUnitStatus3LHistory_(nopol)).forEach(item => {
    const sequence = parseCalibrationNumber_(item.calibrationNumber);
    if (!sequence) return;
    if (!statusTimelineByCalibration[sequence]) statusTimelineByCalibration[sequence] = [];
    statusTimelineByCalibration[sequence].push(item);
  });
  Object.keys(statusTimelineByCalibration).forEach(sequence => {
    statusByCalibration[sequence] = resolveCalibrationStatusSnapshot_(statusTimelineByCalibration[sequence]);
  });
  // Baca base record + FINAL_SNAPSHOT sekaligus. FINAL_SNAPSHOT adalah source of truth
  // untuk siklus yang sudah DONE dan tidak boleh dikalahkan oleh event WAITING/PROCESS
  // yang tercatat sesudahnya akibat data legacy/repair lama.
  const allHistoryRows = readRecentHistoryRows_(sheet, map['NOMOR POLISI'], nopol, 300);
  const trustedStatusIndex = buildTrustedStatus3LStateIndexFromSummaryRows_(nopol, status3LHistory || []);
  const finalSnapshotByCalibration = buildFinalCalibrationSnapshotMap_(allHistoryRows, map, trustedStatusIndex);
  Object.keys(finalSnapshotByCalibration).forEach(sequence => {
    const final = finalSnapshotByCalibration[sequence];
    const evidence = statusByCalibration[sequence] || null;
    if (!isResolvedStatus3LActor_(final.completedBy) && evidence &&
        normalizeLookupText_(evidence.status3L) === 'DONE' && isResolvedStatus3LActor_(evidence.updatedBy)) {
      final.completedBy = evidence.updatedBy;
      final.updatedBy = evidence.updatedBy;
    }
  });
  const baseRows = allHistoryRows
    .filter(row => !map['RECORD TYPE'] ||
      normalizeLookupText_(row[map['RECORD TYPE'] - 1]) !== 'FINAL_SNAPSHOT')
    .slice(0, 30);
  const history = baseRows.map(row => ({
      updatedAt: row[at('TIMESTAMP')] instanceof Date ? row[at('TIMESTAMP')].toISOString() : String(row[at('TIMESTAMP')] || ''),
      updatedBy: status3LActorDisplay_(String(row[at('USER EMAIL')] || '')),
      calibrationNumber: parseCalibrationNumber_(row[at('KALIBRASI KE')]),
      kmDate: unitKmDate_(row[at('TANGGAL KM')]),
      statusUnit: row[at('STATUS UNIT')],
      kmOdometer: row[at('KM ODOMETER')],
      note: String(row[at('CATATAN')] || ''),
      kmPhotoId: historyOdometerPhotoId_(row, map),
      kmPhotoAddedAt: row[at('FOTO DITAMBAHKAN AT')] instanceof Date ? row[at('FOTO DITAMBAHKAN AT')].toISOString() : String(row[at('FOTO DITAMBAHKAN AT')] || ''),
      kmPhotoAddedBy: String(row[at('FOTO DIUNGGAH OLEH')] || row[at('USER EMAIL')] || '')
    })).map(item => {
      const sequence = parseCalibrationNumber_(item.calibrationNumber);
      const finalSnapshot = finalSnapshotByCalibration[sequence] || null;
      const statusEvent = statusByCalibration[sequence] || null;
      const isCurrentCalibration = currentState &&
        parseCalibrationNumber_(currentState.calibrationNumber) === sequence;
      const currentStateEvent = isCurrentCalibration ? {
        status3L: String(currentState.status3L || ''),
        updatedAt: currentState.status3LUpdatedAt || currentState.kmUpdatedAt || '',
        updatedBy: status3LActorDisplay_(currentState.status3LUpdatedBy) || KM_STATUS_3L_SHEET_FALLBACK,
        note: currentState.status3LNote || '',
        source: 'MASTER_UNIT',
        reason: 'CURRENT_STATE'
      } : null;
      // Siklus aktif harus mengikuti primary state MASTER_UNIT. STATUS_3L_HISTORY
      // adalah ledger audit dan dapat tertinggal bila perubahan dibuat automation/API
      // yang tidak memicu onEdit. Siklus historis tetap mengikuti ledger/final snapshot.
      let currentStateUsable = Boolean(currentStateEvent);
      if (currentStateUsable && cleanText_(currentStateEvent.status3L)) {
        try { resolveStatus3L_(currentStateEvent.status3L); }
        catch (error) { currentStateUsable = false; }
      }
      const effectiveStatus = finalSnapshot ||
        (isCurrentCalibration && currentStateUsable ? currentStateEvent : null) ||
        statusEvent;
      item.status3L = effectiveStatus ? effectiveStatus.status3L : '';
      item.status3LUpdatedAt = effectiveStatus ? effectiveStatus.updatedAt : '';
      item.status3LUpdatedBy = effectiveStatus ? effectiveStatus.updatedBy : '';
      item.status3LNote = effectiveStatus ? effectiveStatus.note : '';
      item.status3LSource = finalSnapshot ? 'FINAL_SNAPSHOT' :
        (isCurrentCalibration && currentStateUsable ? 'MASTER_UNIT' :
          (statusEvent ? 'STATUS_3L_HISTORY' : ''));
      item.status3LTimeline = calibrationStatusTimelineForSummary_(
        statusTimelineByCalibration[sequence] || []
      );
      item.finalSnapshot = Boolean(finalSnapshot);
      item.finalSnapshotMissing = Boolean(!finalSnapshot && isCurrentCalibration &&
        normalizeLookupText_(item.status3L) === 'DONE');
      item.lifecycleStatus = (finalSnapshot || normalizeLookupText_(item.status3L) === 'DONE')
        ? 'LOCKED' : 'ACTIVE';
      item.completedAt = finalSnapshot ? finalSnapshot.completedAt :
        (normalizeLookupText_(item.status3L) === 'DONE' ? item.status3LUpdatedAt : '');
      item.completedBy = finalSnapshot ? finalSnapshot.completedBy :
        (normalizeLookupText_(item.status3L) === 'DONE' ? item.status3LUpdatedBy : '');
      return item;
    });
  const latestCalibration = history.reduce((maximum, item) =>
    Math.max(maximum, parseCalibrationNumber_(item.calibrationNumber)), 0);
  history.forEach(item => {
    item.historyLocked = Boolean(item.finalSnapshot) ||
      parseCalibrationNumber_(item.calibrationNumber) < latestCalibration ||
      normalizeLookupText_(item.status3L) === 'DONE';
  });
  return history;
}

function getMasterUnitStatus3LHistory_(nopol) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNIT_STATUS_3L_LOG_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  if (!historyHeadersReady_(sheet, UNIT_STATUS_3L_LOG_HEADERS)) return [];
  const map = getHistoryHeaderMap_(sheet, UNIT_STATUS_3L_LOG_HEADERS);
  const at = header => map[header] - 1;
  return readRecentHistoryRows_(sheet, map['NOMOR POLISI'], nopol, 100).map(row => {
    const updatedBy = status3LActorDisplay_(String(row[at('USER EMAIL')] || ''));
    const status3L = String(row[at('STATUS BARU')] || '');
    const source = String(row[at('SOURCE')] || '');
    const reason = map['REASON'] ? String(row[at('REASON')] || '') : '';
    return {
      updatedAt: row[at('TIMESTAMP')] instanceof Date
        ? row[at('TIMESTAMP')].toISOString() : String(row[at('TIMESTAMP')] || ''),
      updatedBy: updatedBy,
      calibrationNumber: parseCalibrationNumber_(row[at('KALIBRASI KE')]),
      previousStatus: String(row[at('STATUS SEBELUM')] || ''),
      status3L: status3L,
      note: String(row[at('CATATAN')] || ''),
      source: source,
      reason: reason,
      previousStatusUnit: String(row[at('STATUS UNIT SEBELUM')] || ''),
      statusUnit: String(row[at('STATUS UNIT SESUDAH')] || ''),
      trusted: isTrustedStatus3LEvent_(updatedBy, source, reason, status3L),
      weakRepair: isWeakStatus3LRepairReason_(reason) || normalizeLookupText_(source) === 'SYSTEM_REPAIR'
    };
  });
}

function getMasterUnitStatusHistory(nopol, sessionToken) {
  const actor = requirePermission_('VIEW_MASTER_UNIT', '', { token: sessionToken });
  assertUnitScope_(actor, nopol);
  return getMasterUnitStatusHistory_(nopol);
}

function getMasterUnitStatusHistory_(nopol) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNIT_STATUS_LOG_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  if (!historyHeadersReady_(sheet, UNIT_STATUS_LOG_HEADERS)) return [];
  const map = getHistoryHeaderMap_(sheet, UNIT_STATUS_LOG_HEADERS);
  const at = header => map[header] - 1;
  return readRecentHistoryRows_(sheet, map['NOMOR POLISI'], nopol, 30).map(row => ({
      updatedAt: row[at('TIMESTAMP')] instanceof Date ? row[at('TIMESTAMP')].toISOString() : String(row[at('TIMESTAMP')] || ''),
      updatedBy: String(row[at('USER EMAIL')] || ''),
      calibrationNumber: parseCalibrationNumber_(row[at('KALIBRASI KE')]),
      previousStatus: String(row[at('STATUS SEBELUM')] || ''),
      statusUnit: String(row[at('STATUS BARU')] || ''),
      note: String(row[at('CATATAN')] || ''),
      source: String(row[at('SOURCE')] || ''),
      reason: map['REASON'] ? String(row[at('REASON')] || '') : ''
    }));
}

function getMasterUnitDetail_(nopol, target) {
  const record = target === undefined ? readMasterUnitSummaryTarget_(nopol) : target;
  if (!record) return [];
  return record.headers.map((header, index) => ({
    label: cleanText_(header),
    value: formatMasterUnitDetailValue_(header, record.row[index])
  }))
    .filter(field => field.label && field.value);
}


/* ==========================================================================\n * CALIBRATION ANALYTICS / TREND CENTER v29\n * Read-only aggregation dari MASTER_UNIT + HISTORY_KALIBRASI_KM +\n * STATUS_3L_HISTORY. Tidak mengubah workflow operasional maupun history.\n * ========================================================================== */


/** Odometer photo validation, private access, relational audit, and repair. */
const ODOMETER_PHOTO_MAX_BYTES = 5 * 1024 * 1024;
const ODOMETER_PHOTO_FOLDER_ID = '1uccvy7xZ2o868GglUvER5cYLPadwKug7';

function validateOdometerPhoto_(photo) {
  if (photo == null) throw new Error('Foto odometer wajib diunggah untuk setiap pembaruan KM.');
  if (typeof photo !== 'object' || typeof photo.data !== 'string' ||
      !['image/jpeg', 'image/png', 'image/webp'].includes(photo.mimeType)) {
    throw new Error('Foto harus JPEG, PNG, atau WebP.');
  }
  if (!photo.data.length || photo.data.length > Math.ceil(ODOMETER_PHOTO_MAX_BYTES / 3) * 4 ||
      !/^[A-Za-z0-9+/]+={0,2}$/.test(photo.data) || photo.data.length % 4 !== 0) {
    throw new Error('Data foto tidak valid atau melebihi 5 MB.');
  }
  const bytes = Utilities.base64Decode(photo.data);
  const b = bytes.slice(0, 12).map(value => value & 255);
  const valid = photo.mimeType === 'image/jpeg' ? b[0] === 255 && b[1] === 216 && b[2] === 255
    : photo.mimeType === 'image/png' ? [137,80,78,71,13,10,26,10].every((v,i) => b[i] === v)
    : [82,73,70,70].every((v,i) => b[i] === v) && [87,69,66,80].every((v,i) => b[i+8] === v);
  if (!valid || bytes.length > ODOMETER_PHOTO_MAX_BYTES) throw new Error('Isi foto tidak sesuai format atau melebihi 5 MB.');
  return { bytes: bytes, mimeType: photo.mimeType };
}

function saveOdometerPhoto_(photo, nopol) {
  const folder = DriveApp.getFolderById(ODOMETER_PHOTO_FOLDER_ID);
  const extension = { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp' }[photo.mimeType];
  return folder.createFile(Utilities.newBlob(photo.bytes, photo.mimeType,
    normalizeNopol_(nopol) + '-' + Utilities.getUuid() + '.' + extension));
}


/**
 * v17 — Photo integrity recovery.
 *
 * MASTER_UNIT menyimpan pointer foto aktif pada FOTO ODOMETER ID, sedangkan
 * HISTORY_KALIBRASI_KM adalah ledger append-only yang menyimpan pointer foto
 * ketika kalibrasi dibuat. Pada data legacy tertentu file Drive dan ledger
 * tersedia, tetapi pointer pada MASTER_UNIT kosong. Kondisi tersebut bukan
 * "tanpa foto" dan tidak boleh menurunkan Data Quality sebagai missing proof.
 */
function normalizeOdometerPhotoId_(value) {
  const text = cleanText_(value).replace(/^'/, '');
  if (/^[A-Za-z0-9_-]{10,200}$/.test(text)) return text;
  // ID Drive case-sensitive. Terima hanya URL Google yang dikenal, bukan
  // mengambil sembarang token dari URL atau menormalisasi huruf besar/kecil.
  if (!/^https:\/\/(?:drive|docs)\.google\.com\//i.test(text)) return '';
  const match = text.match(/\/d\/([A-Za-z0-9_-]{10,200})(?:[/?#]|$)/) ||
    text.match(/[?&]id=([A-Za-z0-9_-]{10,200})(?:[&#]|$)/);
  return match ? match[1] : '';
}

function historyOdometerPhotoId_(row, map) {
  const after = map['DATA SETELAH'] ? safeParseHistoryJson_(row[map['DATA SETELAH'] - 1]) : {};
  const sequence = parseCalibrationNumber_(getRowValue_(row, map, 'KALIBRASI KE'));
  const afterSequence = parseCalibrationNumber_(after.calibrationNumber);
  // DATA SEBELUM adalah siklus sebelumnya dan tidak boleh dipakai sebagai bukti.
  const snapshotId = (!afterSequence || afterSequence === sequence)
    ? normalizeOdometerPhotoId_(after.kmPhotoId) : '';
  return snapshotId || normalizeOdometerPhotoId_(getRowValue_(row, map, 'FOTO ODOMETER ID'));
}

function addOdometerPhotoRecoveryCandidate_(bucket, key, candidate) {
  if (!key || !candidate || !candidate.photoId) return;
  if (!bucket[key]) bucket[key] = Object.create(null);
  const current = bucket[key][candidate.photoId];
  // Simpan metadata kandidat yang paling lengkap. ID yang sama dapat muncul
  // kembali pada FINAL_SNAPSHOT dan CALIBRATION_OPENED.
  const score = item => !item ? -1 :
    (normalizeLookupText_(item.recordType) === 'CALIBRATION_OPENED' ? 8 : 0) +
    (normalizeLookupText_(item.reason) === 'CALIBRATION_CREATED' ? 4 : 0) +
    (item.addedAt ? 2 : 0) + (item.addedBy ? 1 : 0);
  if (!current || score(candidate) > score(current)) bucket[key][candidate.photoId] = candidate;
}

function odometerPhotoCandidateTime_(candidate) {
  if (!candidate) return 0;
  const raw = candidate.timestamp || candidate.addedAt || '';
  const date = raw instanceof Date ? raw : parseDateTimeInput_(raw);
  return date instanceof Date && !isNaN(date.getTime()) ? date.getTime() : 0;
}

function latestOdometerPhotoCandidate_(current, candidate) {
  if (!current) return candidate;
  return odometerPhotoCandidateTime_(candidate) >= odometerPhotoCandidateTime_(current) ? candidate : current;
}

function odometerPhotoDateValue_(value) {
  if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
  if (value === '' || value == null) return null;
  const parsed = parseDateTimeInput_(value);
  return parsed instanceof Date && !isNaN(parsed.getTime()) ? parsed : null;
}

function odometerPhotoActorKey_(value) {
  return normalizeLookupText_(cleanText_(value));
}

/** Resolve metadata for the exact active photo pointer, never from another ID. */
function odometerPhotoCandidateForId_(index, nopol, calibrationNumber, photoId, kmDate) {
  const normalizedId = normalizeOdometerPhotoId_(photoId);
  if (!index || !normalizedId) return null;
  const exactKey = analyticsUnitKey_(nopol, calibrationNumber);
  const exact = exactKey && index.byCalibration && index.byCalibration[exactKey]
    ? index.byCalibration[exactKey][normalizedId] : null;
  if (exact) return exact;
  const dateKey = dateOnlyKey_(kmDate);
  const fallbackKey = normalizeNopol_(nopol) + '|' + dateKey;
  return dateKey && index.byPlateDate && index.byPlateDate[fallbackKey]
    ? index.byPlateDate[fallbackKey][normalizedId] || null : null;
}

function compareOdometerPhotoMetadata_(addedAt, addedBy, candidate) {
  const actualAt = odometerPhotoDateValue_(addedAt);
  const expectedAt = odometerPhotoDateValue_(candidate && candidate.addedAt);
  const actualBy = cleanText_(addedBy);
  const expectedBy = cleanText_(candidate && candidate.addedBy);
  const now = Date.now();
  return {
    missingAt: addedAt === '' || addedAt == null,
    invalidAt: addedAt !== '' && addedAt != null && !actualAt,
    futureAt: Boolean(actualAt && actualAt.getTime() > now + 5 * 60 * 1000),
    mismatchAt: Boolean(actualAt && expectedAt && Math.abs(actualAt.getTime() - expectedAt.getTime()) > 60 * 1000),
    missingBy: !actualBy,
    mismatchBy: Boolean(actualBy && expectedBy && odometerPhotoActorKey_(actualBy) !== odometerPhotoActorKey_(expectedBy)),
    expectedAt: expectedAt,
    expectedBy: expectedBy,
    recoverableAt: Boolean(expectedAt),
    recoverableBy: Boolean(expectedBy)
  };
}

function buildOdometerPhotoRecoveryIndex_(spreadsheet, historySource) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const history = getUnitKmHistorySheet_(ss);
  const result = {
    schemaReady: false,
    historySheet: history ? history.getName() : '',
    historyRows: history ? Math.max(0, history.getLastRow() - 1) : 0,
    rowsWithPhotoId: 0,
    byCalibration: Object.create(null),
    byPlateDate: Object.create(null),
    openedByCalibration: Object.create(null),
    canonicalByCalibration: Object.create(null),
    rawConflictCount: 0,
    historyMetadata: {
      rowsWithPhotoId: 0,
      missingAddedAt: 0,
      missingAddedBy: 0,
      invalidAddedAt: 0,
      futureAddedAt: 0,
      samples: []
    }
  };
  if (!history || history.getLastRow() < 2 || history.getLastColumn() < 1) return result;

  const map = historySource ? historySource.rawMap : getHeaderMap_(history);
  const required = ['NOMOR POLISI', 'FOTO ODOMETER ID'];
  if (required.some(header => !map[header])) return result;
  result.schemaReady = true;
  const rows = historySource ? historySource.rows : history.getRange(2, 1, history.getLastRow() - 1, history.getLastColumn()).getValues();
  rows.forEach(row => {
    const nopol = normalizeDisplayNopol_(row[map['NOMOR POLISI'] - 1]);
    const photoId = historyOdometerPhotoId_(row, map);
    if (!nopol || !photoId) return;
    result.rowsWithPhotoId++;
    const calibrationNumber = map['KALIBRASI KE']
      ? parseCalibrationNumber_(row[map['KALIBRASI KE'] - 1]) : 0;
    const kmDate = map['TANGGAL KM'] ? dateOnlyKey_(row[map['TANGGAL KM'] - 1]) : '';
    const timestamp = map['TIMESTAMP'] ? row[map['TIMESTAMP'] - 1] : '';
    const directAddedAt = map['FOTO DITAMBAHKAN AT'] ? row[map['FOTO DITAMBAHKAN AT'] - 1] : '';
    const directAddedBy = map['FOTO DIUNGGAH OLEH'] ? displayValue_(row[map['FOTO DIUNGGAH OLEH'] - 1]) : '';
    const userEmail = map['USER EMAIL'] ? displayValue_(row[map['USER EMAIL'] - 1]) : '';
    const recordType = map['RECORD TYPE'] ? displayValue_(row[map['RECORD TYPE'] - 1]) : '';
    const reason = map['REASON'] ? displayValue_(row[map['REASON'] - 1]) : '';
    const recordKey = normalizeLookupText_(recordType);
    const reasonKey = normalizeLookupText_(reason);
    // TIMESTAMP/USER EMAIL hanya ekuivalen dengan metadata upload pada event
    // pembukaan kalibrasi (atau row legacy tanpa lifecycle type). FINAL_SNAPSHOT
    // dan PHOTO_RECONCILIATION memiliki aktor/waktu berbeda dan tidak boleh
    // dipakai untuk menebak siapa/kapan foto awal diunggah.
    const isUploadEvidence = recordKey === 'CALIBRATION_OPENED' ||
      reasonKey === 'CALIBRATION_CREATED' || !recordKey;
    const candidate = {
      photoId: photoId,
      nopol: nopol,
      calibrationNumber: calibrationNumber,
      kmDate: kmDate,
      timestamp: timestamp,
      addedAt: directAddedAt || (isUploadEvidence ? timestamp : ''),
      addedBy: directAddedBy || (isUploadEvidence ? userEmail : ''),
      directAddedAt: directAddedAt,
      directAddedBy: directAddedBy,
      userEmail: userEmail,
      recordType: recordType,
      reason: reason,
      source: history.getName()
    };
    const metadata = result.historyMetadata;
    metadata.rowsWithPhotoId++;
    const addedAtDate = odometerPhotoDateValue_(directAddedAt);
    const timestampDate = odometerPhotoDateValue_(timestamp);
    const issueCodes = [];
    if (!directAddedAt) { metadata.missingAddedAt++; issueCodes.push('MISSING_ADDED_AT'); }
    else if (!addedAtDate) { metadata.invalidAddedAt++; issueCodes.push('INVALID_ADDED_AT'); }
    else if (addedAtDate.getTime() > Date.now() + 5 * 60 * 1000) {
      metadata.futureAddedAt++; issueCodes.push('FUTURE_ADDED_AT');
    }
    if (!cleanText_(directAddedBy)) { metadata.missingAddedBy++; issueCodes.push('MISSING_ADDED_BY'); }
    if (issueCodes.length && metadata.samples.length < 30) {
      metadata.samples.push({
        nopol: nopol,
        calibrationNumber: calibrationNumber,
        photoId: photoId,
        timestamp: timestampDate ? timestampDate.toISOString() : displayValue_(timestamp),
        issues: issueCodes
      });
    }
    const calibrationKey = analyticsUnitKey_(nopol, calibrationNumber);
    addOdometerPhotoRecoveryCandidate_(result.byCalibration, calibrationKey, candidate);

    if (calibrationKey && recordKey === 'CALIBRATION_OPENED' && reasonKey !== 'LEGACY_MASTER_SNAPSHOT') {
      addOdometerPhotoRecoveryCandidate_(result.openedByCalibration, calibrationKey, candidate);
    }
    if (calibrationKey && reasonKey === 'PHOTO_POINTER_RECONCILED_FROM_SUBMIT_EVIDENCE') {
      // Only reconciliation explicitly anchored to submit evidence is canonical.
      // Legacy FROM_MASTER repair rows remain audit history but can never become
      // the source of truth again.
      result.canonicalByCalibration[calibrationKey] = latestOdometerPhotoCandidate_(
        result.canonicalByCalibration[calibrationKey], candidate
      );
    }

    const plateDateKey = normalizeNopol_(nopol) + '|' + kmDate;
    if (kmDate) addOdometerPhotoRecoveryCandidate_(result.byPlateDate, plateDateKey, candidate);
  });

  Object.keys(result.byCalibration).forEach(key => {
    if (Object.keys(result.byCalibration[key] || {}).length > 1) result.rawConflictCount++;
  });
  return result;
}

function resolveOdometerPhotoRecoveryCandidate_(index, nopol, calibrationNumber, kmDate, submitter) {
  const exactKey = analyticsUnitKey_(nopol, calibrationNumber);

  let opened = exactKey && index && index.openedByCalibration && index.openedByCalibration[exactKey]
    ? Object.keys(index.openedByCalibration[exactKey]).map(id => index.openedByCalibration[exactKey][id]) : [];

  // Bukti submit harus lebih kuat daripada PHOTO_RECONCILIATION hasil repair.
  // Repair lama dapat salah ketika MASTER_UNIT sudah membawa pointer unit lain;
  // karena itu canonical repair tidak boleh mengalahkan CALIBRATION_OPENED asli.
  const submitterKey = odometerPhotoActorKey_(submitter);
  if (opened.length && submitterKey) {
    const submitted = opened.filter(candidate =>
      odometerPhotoActorKey_(candidate && (candidate.userEmail || candidate.addedBy)) === submitterKey);
    if (submitted.length === 1) {
      return { status: 'RECOVERABLE', candidate: submitted[0], match: 'SUBMITTER_CALIBRATION_OPENED', canonical: true };
    }
    if (submitted.length > 1) {
      // Legacy migration dapat membuat >1 CALIBRATION_OPENED untuk user yang sama.
      // Pilih bukti yang waktu upload fotonya paling dekat dengan timestamp event submit.
      const ranked = submitted.slice().sort((a, b) => {
        const gap = item => {
          const t = odometerPhotoDateValue_(item && item.timestamp);
          const p = odometerPhotoDateValue_(item && item.addedAt);
          return t && p ? Math.abs(t.getTime() - p.getTime()) : Number.MAX_SAFE_INTEGER;
        };
        const delta = gap(a) - gap(b);
        if (delta) return delta;
        return odometerPhotoCandidateTime_(b) - odometerPhotoCandidateTime_(a);
      });
      if (ranked.length) {
        return { status: 'RECOVERABLE', candidate: ranked[0], match: 'SUBMITTER_CALIBRATION_OPENED_RANKED', canonical: true };
      }
    }
  }

  // CALIBRATION_OPENED adalah bukti kuat karena ditulis pada transaksi yang sama
  // dengan MASTER_UNIT. Jika hanya ada satu ID pada pembukaan siklus, FINAL_SNAPSHOT
  // legacy yang membawa pointer lama tidak boleh membuat false positive ambiguous.
  if (opened.length === 1) {
    return { status: 'RECOVERABLE', candidate: opened[0], match: 'CALIBRATION_OPENED', canonical: true };
  }
  if (opened.length > 1) {
    return { status: 'AMBIGUOUS', candidates: opened, match: 'CALIBRATION_OPENED', canonical: false };
  }

  const explicitCanonical = exactKey && index && index.canonicalByCalibration
    ? index.canonicalByCalibration[exactKey] : null;
  if (explicitCanonical && explicitCanonical.photoId) {
    return { status: 'RECOVERABLE', candidate: explicitCanonical, match: 'PHOTO_RECONCILIATION', canonical: true };
  }

  let candidates = exactKey && index && index.byCalibration && index.byCalibration[exactKey]
    ? Object.keys(index.byCalibration[exactKey]).map(id => index.byCalibration[exactKey][id]) : [];
  if (candidates.length === 1) return { status: 'RECOVERABLE', candidate: candidates[0], match: 'CALIBRATION', canonical: true };
  if (candidates.length > 1) return { status: 'AMBIGUOUS', candidates: candidates, match: 'CALIBRATION', canonical: false };

  const dateKey = dateOnlyKey_(kmDate);
  const fallbackKey = normalizeNopol_(nopol) + '|' + dateKey;
  candidates = dateKey && index && index.byPlateDate && index.byPlateDate[fallbackKey]
    ? Object.keys(index.byPlateDate[fallbackKey]).map(id => index.byPlateDate[fallbackKey][id]) : [];
  candidates = candidates.filter(candidate => !candidate.calibrationNumber ||
    candidate.calibrationNumber === parseCalibrationNumber_(calibrationNumber));
  if (candidates.length === 1) return { status: 'RECOVERABLE', candidate: candidates[0], match: 'NOPOL_DATE', canonical: true };
  if (candidates.length > 1) return { status: 'AMBIGUOUS', candidates: candidates, match: 'NOPOL_DATE', canonical: false };
  return { status: 'MISSING', candidates: [], match: '', canonical: false };
}

function auditOdometerPhotoIntegrity_(masterSheet, indexes, recoveryIndex) {
  const audit = {
    totalCalibrated: 0,
    masterPhotoIdPresent: 0,
    invalidPhotoId: 0,
    normalizedIdNeeded: 0,
    recoverable: 0,
    mismatch: 0,
    ambiguous: 0,
    reconciled: 0,
    rawHistoryConflicts: Number(recoveryIndex && recoveryIndex.rawConflictCount || 0),
    missingEvidence: 0,
    orphanMetadata: 0,
    metadataMissingAt: 0,
    metadataMissingBy: 0,
    metadataInvalidAt: 0,
    metadataFutureAt: 0,
    metadataMismatchAt: 0,
    metadataMismatchBy: 0,
    metadataRepairable: 0,
    metadataUnresolved: 0,
    historyMetadata: recoveryIndex && recoveryIndex.historyMetadata || {},
    samples: {
      recoverable: [], mismatch: [], ambiguous: [], reconciled: [], missingEvidence: [],
      invalidPhotoId: [], orphanMetadata: [], metadata: []
    }
  };
  if (!masterSheet || masterSheet.getLastRow() < 2 || !indexes ||
      indexes.nopol < 0 || indexes.calibrationNumber < 0 || indexes.kmPhotoId < 0) return audit;
  const rows = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, masterSheet.getLastColumn()).getValues();
  rows.forEach(row => {
    const nopol = normalizeDisplayNopol_(row[indexes.nopol]);
    const calibration = parseCalibrationNumber_(row[indexes.calibrationNumber]);
    if (!nopol || !calibration) return;
    audit.totalCalibrated++;
    const rawMasterPhotoId = cleanText_(row[indexes.kmPhotoId]);
    const masterPhotoId = normalizeOdometerPhotoId_(rawMasterPhotoId);
    const masterAddedAt = indexes.kmPhotoAddedAt >= 0 ? row[indexes.kmPhotoAddedAt] : '';
    const masterAddedBy = indexes.kmPhotoAddedBy >= 0 ? row[indexes.kmPhotoAddedBy] : '';
    const kmDate = indexes.kmDate >= 0 ? row[indexes.kmDate] : '';
    const submitter = indexes.kmUpdatedBy >= 0 ? cleanText_(row[indexes.kmUpdatedBy]) : '';
    const resolved = resolveOdometerPhotoRecoveryCandidate_(recoveryIndex, nopol, calibration, kmDate, submitter);
    const sample = { nopol: nopol, calibrationNumber: calibration, kmDate: formatDateOnly_(kmDate) };
    const exactKey = analyticsUnitKey_(nopol, calibration);
    const rawIds = exactKey && recoveryIndex && recoveryIndex.byCalibration && recoveryIndex.byCalibration[exactKey]
      ? Object.keys(recoveryIndex.byCalibration[exactKey]) : [];
    if (rawMasterPhotoId && !masterPhotoId) {
      audit.invalidPhotoId++;
      if (audit.samples.invalidPhotoId.length < 20) audit.samples.invalidPhotoId.push(Object.assign(sample, {
        rawPhotoId: rawMasterPhotoId
      }));
    } else if (masterPhotoId && rawMasterPhotoId !== masterPhotoId) {
      audit.normalizedIdNeeded++;
    }
    if (!masterPhotoId && (masterAddedAt || cleanText_(masterAddedBy))) {
      audit.orphanMetadata++;
      if (audit.samples.orphanMetadata.length < 20) audit.samples.orphanMetadata.push(Object.assign(sample, {
        addedAt: displayValue_(masterAddedAt), addedBy: displayValue_(masterAddedBy)
      }));
    }
    if (masterPhotoId) {
      audit.masterPhotoIdPresent++;
      const metadataCandidate = odometerPhotoCandidateForId_(
        recoveryIndex, nopol, calibration, masterPhotoId, kmDate
      );
      const metadata = compareOdometerPhotoMetadata_(masterAddedAt, masterAddedBy, metadataCandidate);
      const metadataIssues = [];
      if (metadata.missingAt) { audit.metadataMissingAt++; metadataIssues.push('MISSING_ADDED_AT'); }
      if (metadata.missingBy) { audit.metadataMissingBy++; metadataIssues.push('MISSING_ADDED_BY'); }
      if (metadata.invalidAt) { audit.metadataInvalidAt++; metadataIssues.push('INVALID_ADDED_AT'); }
      if (metadata.futureAt) { audit.metadataFutureAt++; metadataIssues.push('FUTURE_ADDED_AT'); }
      if (metadata.mismatchAt) { audit.metadataMismatchAt++; metadataIssues.push('MISMATCH_ADDED_AT'); }
      if (metadata.mismatchBy) { audit.metadataMismatchBy++; metadataIssues.push('MISMATCH_ADDED_BY'); }
      if (metadataIssues.length) {
        const canRepair = (!metadata.missingAt && !metadata.invalidAt && !metadata.futureAt && !metadata.mismatchAt || metadata.recoverableAt) &&
          (!metadata.missingBy && !metadata.mismatchBy || metadata.recoverableBy);
        if (canRepair) audit.metadataRepairable++; else audit.metadataUnresolved++;
        if (audit.samples.metadata.length < 30) audit.samples.metadata.push(Object.assign(sample, {
          photoId: masterPhotoId,
          issues: metadataIssues,
          historyMatch: metadataCandidate ? metadataCandidate.recordType || metadataCandidate.reason || 'HISTORY' : '',
          repairable: canRepair
        }));
      }
      if (resolved.status === 'RECOVERABLE') {
        const canonicalId = normalizeOdometerPhotoId_(resolved.candidate && resolved.candidate.photoId);
        if (canonicalId && canonicalId !== masterPhotoId) {
          audit.mismatch++;
          if (audit.samples.mismatch.length < 20) audit.samples.mismatch.push(Object.assign(sample, {
            masterPhotoId: masterPhotoId, canonicalPhotoId: canonicalId, match: resolved.match, historyPhotoIds: rawIds
          }));
        } else if (resolved.match === 'PHOTO_RECONCILIATION' && rawIds.length > 1) {
          audit.reconciled++;
          if (audit.samples.reconciled.length < 20) audit.samples.reconciled.push(Object.assign(sample, {
            masterPhotoId: masterPhotoId, canonicalPhotoId: canonicalId, historyPhotoIds: rawIds
          }));
        }
      } else if (resolved.status === 'AMBIGUOUS') {
        audit.ambiguous++;
        if (audit.samples.ambiguous.length < 20) audit.samples.ambiguous.push(Object.assign(sample, {
          masterPhotoId: masterPhotoId,
          candidateCount: resolved.candidates.length,
          historyPhotoIds: resolved.candidates.map(item => item.photoId)
        }));
      }
      return;
    }
    if (resolved.status === 'RECOVERABLE') {
      audit.recoverable++;
      if (audit.samples.recoverable.length < 20) audit.samples.recoverable.push(Object.assign(sample, {
        photoId: resolved.candidate.photoId, match: resolved.match
      }));
    } else if (resolved.status === 'AMBIGUOUS') {
      audit.ambiguous++;
      if (audit.samples.ambiguous.length < 20) audit.samples.ambiguous.push(Object.assign(sample, {
        candidateCount: resolved.candidates.length,
        historyPhotoIds: resolved.candidates.map(item => item.photoId)
      }));
    } else {
      audit.missingEvidence++;
      if (audit.samples.missingEvidence.length < 20) audit.samples.missingEvidence.push(sample);
    }
  });
  return audit;
}

function repairOdometerPhotoIdsFromHistory_(actor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  if (!master || master.getLastRow() < 2) {
    return { success: true, recovered: 0, ambiguous: 0, missingEvidence: 0, scanned: 0 };
  }
  const headers = master.getRange(1, 1, 1, master.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  ['nopol', 'calibrationNumber', 'kmPhotoId'].forEach(field => {
    if (indexes[field] < 0) throw new Error('DATABASE_NOT_READY: kolom ' + field + ' belum tersedia di MASTER_UNIT.');
  });
  const recoveryIndex = buildOdometerPhotoRecoveryIndex_(ss);
  if (!recoveryIndex.schemaReady) {
    throw new Error('DATABASE_NOT_READY: HISTORY_KALIBRASI_KM belum memiliki NOMOR POLISI dan FOTO ODOMETER ID yang dapat dibaca.');
  }

  const lock = acquireScriptLock_(30000, 'sinkronisasi ID foto odometer');
  try {
    const rowCount = master.getLastRow() - 1;
    const rows = master.getRange(2, 1, rowCount, master.getLastColumn()).getValues();
    const photoIds = master.getRange(2, indexes.kmPhotoId + 1, rowCount, 1).getValues();
    const photoAddedAt = indexes.kmPhotoAddedAt >= 0
      ? master.getRange(2, indexes.kmPhotoAddedAt + 1, rowCount, 1).getValues() : null;
    const photoAddedBy = indexes.kmPhotoAddedBy >= 0
      ? master.getRange(2, indexes.kmPhotoAddedBy + 1, rowCount, 1).getValues() : null;
    let recovered = 0, ambiguous = 0, missingEvidence = 0, scanned = 0;
    const recoveredSamples = [];

    rows.forEach((row, offset) => {
      const nopol = normalizeDisplayNopol_(row[indexes.nopol]);
      const calibration = parseCalibrationNumber_(row[indexes.calibrationNumber]);
      if (!nopol || !calibration || normalizeOdometerPhotoId_(photoIds[offset][0])) return;
      scanned++;
      const kmDate = indexes.kmDate >= 0 ? row[indexes.kmDate] : '';
      const submitter = indexes.kmUpdatedBy >= 0 ? row[indexes.kmUpdatedBy] : '';
      const resolved = resolveOdometerPhotoRecoveryCandidate_(recoveryIndex, nopol, calibration, kmDate, submitter);
      if (resolved.status === 'AMBIGUOUS') { ambiguous++; return; }
      if (resolved.status !== 'RECOVERABLE') { missingEvidence++; return; }
      const candidate = resolved.candidate;
      photoIds[offset][0] = safeSheetText_(candidate.photoId);
      // Ketika pointer MASTER_UNIT kosong, metadata lama adalah orphan. Selaraskan
      // keduanya dengan evidence pada row history yang menghasilkan pointer.
      if (photoAddedAt) photoAddedAt[offset][0] = candidate.addedAt || candidate.timestamp || '';
      if (photoAddedBy) photoAddedBy[offset][0] = safeSheetText_(candidate.addedBy || '');
      recovered++;
      if (recoveredSamples.length < 30) recoveredSamples.push({
        nopol: nopol, calibrationNumber: calibration, photoId: candidate.photoId,
        match: resolved.match, source: candidate.source
      });
    });

    if (recovered) {
      master.getRange(2, indexes.kmPhotoId + 1, rowCount, 1).setValues(photoIds);
      if (photoAddedAt) master.getRange(2, indexes.kmPhotoAddedAt + 1, rowCount, 1).setValues(photoAddedAt);
      if (photoAddedBy) master.getRange(2, indexes.kmPhotoAddedBy + 1, rowCount, 1).setValues(photoAddedBy);
      SpreadsheetApp.flush();
      invalidateAppCaches_();
    }

    const result = {
      success: true,
      scanned: scanned,
      recovered: recovered,
      ambiguous: ambiguous,
      missingEvidence: missingEvidence,
      historyRows: recoveryIndex.historyRows,
      historyRowsWithPhotoId: recoveryIndex.rowsWithPhotoId,
      samples: recoveredSamples
    };
    logSystemActivity_('', '', 'ODOMETER_PHOTO_ID_RECOVERY', 'MASTER_UNIT:FOTO ODOMETER ID', '',
      JSON.stringify({ recovered: recovered, ambiguous: ambiguous, missingEvidence: missingEvidence }),
      'Sinkronisasi pointer foto dari HISTORY_KALIBRASI_KM; file Drive dan history tidak diubah.', actor, true);
    return result;
  } finally {
    lock.releaseLock();
  }
}

function appendPhotoReconciliationRows_(entries) {
  const rows = Array.isArray(entries) ? entries.filter(Boolean) : [];
  if (!rows.length) return { written: 0, skippedExisting: 0 };
  const sheet = ensureUnitKmHistorySheet_(SpreadsheetApp.getActiveSpreadsheet());
  const map = ensureHistoryHeaders_(sheet, UNIT_KM_LOG_HEADERS);
  const latestCanonicalByKey = Object.create(null);
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues().forEach(row => {
      const type = normalizeLookupText_(getRowValue_(row, map, 'RECORD TYPE'));
      const reason = normalizeLookupText_(getRowValue_(row, map, 'REASON'));
      if (type !== 'PHOTO_RECONCILIATION' || reason !== 'PHOTO_POINTER_RECONCILED_FROM_SUBMIT_EVIDENCE') return;
      const key = analyticsUnitKey_(getRowValue_(row, map, 'NOMOR POLISI'), getRowValue_(row, map, 'KALIBRASI KE'));
      const photoId = historyOdometerPhotoId_(row, map);
      const timeRaw = getRowValue_(row, map, 'TIMESTAMP');
      const time = timeRaw instanceof Date ? timeRaw.getTime() : ((parseDateTimeInput_(timeRaw) || new Date(0)).getTime());
      if (!key || !photoId) return;
      const current = latestCanonicalByKey[key];
      if (!current || time >= current.time) latestCanonicalByKey[key] = { photoId: photoId, time: time };
    });
  }
  const pending = rows.filter(entry => {
    const key = analyticsUnitKey_(entry.nopol, entry.calibrationNumber);
    const photoId = normalizeOdometerPhotoId_(entry.photoId);
    if (!key || !photoId) return false;
    const current = latestCanonicalByKey[key];
    if (current && current.photoId === photoId) return false;
    latestCanonicalByKey[key] = { photoId: photoId, time: Date.now() };
    return true;
  });
  if (!pending.length) return { written: 0, skippedExisting: rows.length };

  const now = new Date();
  const values = pending.map(entry => {
    const snapshot = entry.snapshot || {};
    const timestamp = entry.timestamp instanceof Date ? entry.timestamp : now;
    const updateId = Utilities.getUuid();
    return buildHistoryRow_(sheet, map, {
      'UPDATE ID': updateId,
      'HISTORY ID': newCalibrationHistoryId_(timestamp),
      'REQUEST ID': updateId,
      'CALIBRATION SEQUENCE': parseCalibrationNumber_(entry.calibrationNumber),
      'RECORD TYPE': 'PHOTO_RECONCILIATION',
      'LIFECYCLE STATUS': normalizeLookupText_(snapshot.status3L) === 'DONE' ? 'LOCKED' : 'ACTIVE',
      'TIMESTAMP': timestamp,
      'USER EMAIL': safeSheetText_(entry.email || 'SYSTEM_REPAIR'),
      'NOMOR POLISI': safeSheetText_(entry.nopol),
      'KALIBRASI KE': parseCalibrationNumber_(entry.calibrationNumber),
      'TANGGAL KM': snapshot.kmDate ? parseDateInput_(snapshot.kmDate) || snapshot.kmDate : '',
      'STATUS UNIT': safeSheetText_(snapshot.statusUnit),
      'KM ODOMETER': snapshot.kmOdometer == null ? '' : snapshot.kmOdometer,
      'CATATAN': safeSheetText_(entry.note || 'Rekonsiliasi pointer foto canonical dari state aktif MASTER_UNIT; histori lama tidak dihapus.'),
      'FOTO ODOMETER ID': normalizeOdometerPhotoId_(entry.photoId),
      // Jangan mengarang waktu/uploader foto dari waktu/aktor repair. Kolom
      // USER EMAIL + TIMESTAMP sudah menyimpan siapa dan kapan rekonsiliasi dibuat.
      'FOTO DITAMBAHKAN AT': snapshot.kmPhotoAddedAt ? parseDateTimeInput_(snapshot.kmPhotoAddedAt) || snapshot.kmPhotoAddedAt : '',
      'FOTO DIUNGGAH OLEH': safeSheetText_(snapshot.kmPhotoAddedBy || ''),
      'OPERATING POINT': safeSheetText_(entry.operatingPoint),
      'NAMA TOH/WOH': safeSheetText_(entry.toh),
      'STATUS 3L AWAL': safeSheetText_(snapshot.status3L),
      'STATUS 3L UPDATED AT': snapshot.status3LUpdatedAt ? parseDateTimeInput_(snapshot.status3LUpdatedAt) || snapshot.status3LUpdatedAt : '',
      'STATUS 3L UPDATED BY': safeSheetText_(snapshot.status3LUpdatedBy),
      'COMPLETED AT': normalizeLookupText_(snapshot.status3L) === 'DONE' && snapshot.status3LUpdatedAt
        ? (parseDateTimeInput_(snapshot.status3LUpdatedAt) || snapshot.status3LUpdatedAt) : '',
      'COMPLETED BY': normalizeLookupText_(snapshot.status3L) === 'DONE' ? safeSheetText_(snapshot.status3LUpdatedBy) : '',
      'SOURCE': safeSheetText_(entry.source || 'SYSTEM_REPAIR'),
      'REASON': safeSheetText_(entry.reason || 'PHOTO_POINTER_RECONCILED_FROM_SUBMIT_EVIDENCE'),
      'DATA SETELAH': JSON.stringify(snapshot)
    });
  });
  ensureSheetCapacity_(sheet, values.length);
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, sheet.getLastColumn()).setValues(values);
  return { written: values.length, skippedExisting: rows.length - values.length };
}

/**
 * Repair photo integrity (canonical per siklus).
 * - Pointer MASTER_UNIT kosong: isi hanya dari canonical history yang unik.
 * - Pointer MASTER_UNIT tersedia tetapi history legacy berbeda/berkonflik:
 *   jangan rewrite history; tambahkan PHOTO_RECONCILIATION append-only agar
 *   dashboard memiliki satu pointer canonical yang dapat diaudit.
 */
function repairCalibrationPhotoIntegrity_(actor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  if (!master || master.getLastRow() < 2) {
    return { success: true, scanned: 0, recovered: 0, reconciled: 0, ambiguous: 0, missingEvidence: 0 };
  }
  const headers = master.getRange(1, 1, 1, master.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  ['nopol','calibrationNumber','kmPhotoId'].forEach(field => {
    if (indexes[field] < 0) throw new Error('DATABASE_NOT_READY: kolom ' + field + ' belum tersedia di MASTER_UNIT.');
  });

  const lock = acquireScriptLock_(30000, 'repair photo integrity');
  try {
    const rowCount = master.getLastRow() - 1;
    const rows = master.getRange(2, 1, rowCount, master.getLastColumn()).getValues();
    const photoIds = master.getRange(2, indexes.kmPhotoId + 1, rowCount, 1).getValues();
    const photoAddedAt = indexes.kmPhotoAddedAt >= 0
      ? master.getRange(2, indexes.kmPhotoAddedAt + 1, rowCount, 1).getValues() : null;
    const photoAddedBy = indexes.kmPhotoAddedBy >= 0
      ? master.getRange(2, indexes.kmPhotoAddedBy + 1, rowCount, 1).getValues() : null;
    const index = buildOdometerPhotoRecoveryIndex_(ss);
    let scanned = 0, recovered = 0, normalizedIds = 0, ambiguous = 0, missingEvidence = 0;
    let metadataRecovered = 0, metadataCorrected = 0, metadataUnresolved = 0;
    let masterChanged = false;
    const reconciliationEntries = [];
    const samples = [];

    rows.forEach((row, offset) => {
      const nopol = normalizeDisplayNopol_(row[indexes.nopol]);
      const calibration = parseCalibrationNumber_(row[indexes.calibrationNumber]);
      if (!nopol || !calibration) return;
      scanned++;
      const rawMasterPhotoId = cleanText_(photoIds[offset][0]);
      const masterPhotoId = normalizeOdometerPhotoId_(rawMasterPhotoId);
      const kmDate = indexes.kmDate >= 0 ? row[indexes.kmDate] : '';
      const submitter = indexes.kmUpdatedBy >= 0 ? row[indexes.kmUpdatedBy] : '';
      const resolved = resolveOdometerPhotoRecoveryCandidate_(index, nopol, calibration, kmDate, submitter);
      const key = analyticsUnitKey_(nopol, calibration);
      const rawIds = key && index.byCalibration[key] ? Object.keys(index.byCalibration[key]) : [];

      if (!masterPhotoId) {
        if (resolved.status === 'RECOVERABLE') {
          const candidate = resolved.candidate;
          photoIds[offset][0] = safeSheetText_(candidate.photoId);
          if (photoAddedAt) photoAddedAt[offset][0] = candidate.addedAt || candidate.timestamp || '';
          if (photoAddedBy) photoAddedBy[offset][0] = safeSheetText_(candidate.addedBy || '');
          recovered++;
          masterChanged = true;
          if (samples.length < 50) samples.push({ nopol:nopol, calibrationNumber:calibration, action:'RECOVERED_MASTER_POINTER', photoId:candidate.photoId, match:resolved.match });
        } else if (resolved.status === 'AMBIGUOUS') {
          ambiguous++;
        } else {
          missingEvidence++;
        }
        return;
      }

      if (rawMasterPhotoId !== masterPhotoId) {
        photoIds[offset][0] = safeSheetText_(masterPhotoId);
        normalizedIds++;
        masterChanged = true;
      }

      // Jika bukti CALIBRATION_OPENED milik user submit tersedia, bukti submit
      // adalah sumber kebenaran. Jangan mempertahankan pointer MASTER_UNIT yang
      // berbeda karena itulah pola bug SYSTEM_REPAIR yang pernah memindahkan
      // foto antar-unit.
      const resolvedPhotoId = resolved.status === 'RECOVERABLE'
        ? normalizeOdometerPhotoId_(resolved.candidate && resolved.candidate.photoId) : '';
      const resolvedFromSubmit = resolvedPhotoId && (
        resolved.match === 'SUBMITTER_CALIBRATION_OPENED' ||
        resolved.match === 'SUBMITTER_CALIBRATION_OPENED_RANKED' ||
        resolved.match === 'CALIBRATION_OPENED'
      );
      if (resolvedFromSubmit && resolvedPhotoId !== masterPhotoId) {
        const candidate = resolved.candidate;
        photoIds[offset][0] = safeSheetText_(resolvedPhotoId);
        if (photoAddedAt) photoAddedAt[offset][0] = candidate.addedAt || candidate.timestamp || '';
        if (photoAddedBy) photoAddedBy[offset][0] = safeSheetText_(candidate.addedBy || candidate.userEmail || submitter || '');
        masterChanged = true;
        if (samples.length < 50) samples.push({
          nopol:nopol, calibrationNumber:calibration, action:'CORRECTED_MASTER_FROM_SUBMIT_EVIDENCE',
          previousPhotoId:masterPhotoId, photoId:resolvedPhotoId, match:resolved.match
        });
      }

      const effectiveMasterPhotoId = normalizeOdometerPhotoId_(photoIds[offset][0]);

      // Metadata wajib mengikuti ID foto aktif yang sama. Jangan mengambil
      // metadata dari canonical ID lain pada siklus yang konflik.
      const metadataCandidate = odometerPhotoCandidateForId_(index, nopol, calibration, effectiveMasterPhotoId, kmDate);
      const metadata = compareOdometerPhotoMetadata_(
        photoAddedAt ? photoAddedAt[offset][0] : '',
        photoAddedBy ? photoAddedBy[offset][0] : '',
        metadataCandidate
      );
      let metadataChangedOnRow = false;
      let unresolvedOnRow = false;
      if (metadata.missingAt || metadata.invalidAt || metadata.futureAt || metadata.mismatchAt) {
        if (photoAddedAt && metadata.recoverableAt) {
          photoAddedAt[offset][0] = metadata.expectedAt;
          metadataChangedOnRow = true;
          if (metadata.mismatchAt) metadataCorrected++; else metadataRecovered++;
        } else {
          unresolvedOnRow = true;
        }
      }
      if (metadata.missingBy || metadata.mismatchBy) {
        if (photoAddedBy && metadata.recoverableBy) {
          photoAddedBy[offset][0] = safeSheetText_(metadata.expectedBy);
          metadataChangedOnRow = true;
          if (metadata.mismatchBy) metadataCorrected++; else metadataRecovered++;
        } else {
          unresolvedOnRow = true;
        }
      }
      if (unresolvedOnRow) metadataUnresolved++;
      if (metadataChangedOnRow) {
        masterChanged = true;
        if (samples.length < 50) samples.push({
          nopol:nopol, calibrationNumber:calibration, action:'SYNCED_PHOTO_METADATA',
          photoId:effectiveMasterPhotoId, source:metadataCandidate && metadataCandidate.source || ''
        });
      }

      const explicit = index.canonicalByCalibration && index.canonicalByCalibration[key];
      const explicitId = explicit ? normalizeOdometerPhotoId_(explicit.photoId) : '';
      const canonicalId = resolved.status === 'RECOVERABLE' ? normalizeOdometerPhotoId_(resolved.candidate && resolved.candidate.photoId) : '';
      const needsReconciliation = explicitId !== effectiveMasterPhotoId && (
        !canonicalId || canonicalId !== effectiveMasterPhotoId || rawIds.length > 1 || !rawIds.includes(effectiveMasterPhotoId)
      );
      if (!needsReconciliation) return;
      if (!resolvedFromSubmit) {
        // A mismatch without independent CALIBRATION_OPENED evidence is not safe
        // to auto-canonicalize. Keep the ledger intact and surface it for review.
        ambiguous++;
        if (samples.length < 50) samples.push({
          nopol:nopol, calibrationNumber:calibration,
          action:'SKIPPED_UNTRUSTED_MASTER_RECONCILIATION',
          masterPhotoId:effectiveMasterPhotoId, previousHistoryPhotoIds:rawIds
        });
        return;
      }

      const reconciledRow = row.slice();
      reconciledRow[indexes.kmPhotoId] = effectiveMasterPhotoId;
      if (photoAddedAt) reconciledRow[indexes.kmPhotoAddedAt] = photoAddedAt[offset][0];
      if (photoAddedBy) reconciledRow[indexes.kmPhotoAddedBy] = photoAddedBy[offset][0];
      reconciliationEntries.push({
        nopol: nopol,
        calibrationNumber: calibration,
        photoId: effectiveMasterPhotoId,
        operatingPoint: indexes.operatingPoint >= 0 ? displayValue_(row[indexes.operatingPoint]) : '',
        toh: indexes.toh >= 0 ? displayValue_(row[indexes.toh]) : '',
        email: actor && actor.email || 'SYSTEM_REPAIR',
        timestamp: new Date(),
        source: 'SYSTEM_REPAIR',
        reason: 'PHOTO_POINTER_RECONCILED_FROM_SUBMIT_EVIDENCE',
        note: 'Pointer foto dikoreksi berdasarkan CALIBRATION_OPENED milik user submit; MASTER_UNIT tidak dijadikan sumber kebenaran saat berbeda dengan bukti submit.',
        snapshot: unitKmSnapshot_(reconciledRow, indexes)
      });
      if (samples.length < 50) samples.push({ nopol:nopol, calibrationNumber:calibration, action:'APPEND_CANONICAL_RECONCILIATION', masterPhotoId:effectiveMasterPhotoId, previousHistoryPhotoIds:rawIds });
    });

    if (masterChanged) {
      master.getRange(2, indexes.kmPhotoId + 1, rowCount, 1).setValues(photoIds);
      if (photoAddedAt) master.getRange(2, indexes.kmPhotoAddedAt + 1, rowCount, 1).setValues(photoAddedAt);
      if (photoAddedBy) master.getRange(2, indexes.kmPhotoAddedBy + 1, rowCount, 1).setValues(photoAddedBy);
    }
    const reconciliation = appendPhotoReconciliationRows_(reconciliationEntries);
    SpreadsheetApp.flush();
    invalidateAppCaches_();
    const result = {
      success: true,
      scanned: scanned,
      recovered: recovered,
      normalizedIds: normalizedIds,
      metadataRecovered: metadataRecovered,
      metadataCorrected: metadataCorrected,
      metadataUnresolved: metadataUnresolved,
      reconciled: reconciliation.written,
      reconciliationSkippedExisting: reconciliation.skippedExisting,
      ambiguous: ambiguous,
      missingEvidence: missingEvidence,
      samples: samples
    };
    logSystemActivity_('', '', 'CALIBRATION_PHOTO_INTEGRITY_REPAIR', 'MASTER_UNIT/HISTORY_KALIBRASI_KM', '', JSON.stringify(result),
      'Pointer aktif MASTER_UNIT dipertahankan sebagai canonical current-cycle; konflik history direkonsiliasi append-only.', actor, true);
    return result;
  } finally {
    lock.releaseLock();
  }
}

function repairCalibrationPhotoIntegrity(sessionToken) {
  const actor = requirePermission_('ADMIN_SETUP', '', { token: sessionToken });
  return repairCalibrationPhotoIntegrity_(actor);
}

/** Audit read-only dari menu Spreadsheet; tidak mengubah MASTER_UNIT/history/Drive. */
function auditCalibrationPhotoIntegrityFromSheet() {
  requirePermission_('ADMIN_SETUP', '', { source: 'SHEET' });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  if (!master) throw new Error('MASTER_UNIT belum tersedia.');
  const headers = master.getRange(1, 1, 1, master.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  const audit = auditOdometerPhotoIntegrity_(master, indexes, buildOdometerPhotoRecoveryIndex_(ss));
  const history = audit.historyMetadata || {};
  const message = [
    'Siklus terkalibrasi: ' + Number(audit.totalCalibrated || 0),
    'Pointer aktif: ' + Number(audit.masterPhotoIdPresent || 0),
    'ID dapat dipulihkan: ' + Number(audit.recoverable || 0),
    'ID invalid / perlu normalisasi: ' + (Number(audit.invalidPhotoId || 0) + Number(audit.normalizedIdNeeded || 0)),
    'Mismatch / ambigu: ' + (Number(audit.mismatch || 0) + Number(audit.ambiguous || 0)),
    'Metadata MASTER_UNIT perlu perhatian: ' + (
      Number(audit.metadataMissingAt || 0) + Number(audit.metadataMissingBy || 0) +
      Number(audit.metadataInvalidAt || 0) + Number(audit.metadataFutureAt || 0) +
      Number(audit.metadataMismatchAt || 0) + Number(audit.metadataMismatchBy || 0)
    ),
    'Metadata tanpa evidence untuk repair: ' + Number(audit.metadataUnresolved || 0),
    'History tanpa waktu / uploader eksplisit: ' + Number(history.missingAddedAt || 0) + ' / ' + Number(history.missingAddedBy || 0),
    '',
    'Detail dan sample tersedia di Administration > Master Unit > Data Quality.'
  ].join('\n');
  console.log(JSON.stringify(audit));
  SpreadsheetApp.getUi().alert('Audit Relasi Foto (Read-only)', message, SpreadsheetApp.getUi().ButtonSet.OK);
  return audit;
}

function repairCalibrationPhotoIntegrityFromSheet() {
  const actor = requirePermission_('ADMIN_SETUP', '', { source: 'SHEET' });
  return repairCalibrationPhotoIntegrity_(actor);
}

function repairOdometerPhotoIdsFromHistory(sessionToken) {
  const actor = requirePermission_('ADMIN_SETUP', '', { token: sessionToken });
  return repairOdometerPhotoIdsFromHistory_(actor);
}

function repairOdometerPhotoIdsFromHistoryFromSheet() {
  const actor = requirePermission_('ADMIN_SETUP', '', { source: 'SHEET' });
  return repairOdometerPhotoIdsFromHistory_(actor);
}

/** Foto privat dibaca melalui sesi dan scope unit, bukan link publik Drive. */
function getOdometerPhoto(nopol, photoId, sessionToken) {
  const actor = requirePermission_('VIEW_MASTER_UNIT', '', { token: sessionToken });
  assertUnitScope_(actor, nopol);
  photoId = normalizeOdometerPhotoId_(photoId);
  if (!photoId) throw new Error('ID foto tidak valid.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  if (!master || master.getLastRow() < 2) throw new Error('MASTER_UNIT belum tersedia.');
  const headers = master.getRange(1, 1, 1, master.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  if (indexes.nopol < 0 || indexes.kmPhotoId < 0) throw new Error('Kolom foto odometer belum tersedia.');

  const key = normalizeNopol_(nopol);
  let photoAudit = null;
  const rowCount = master.getLastRow() - 1;
  const plates = master.getRange(2, indexes.nopol + 1, rowCount, 1).getDisplayValues();
  const photos = master.getRange(2, indexes.kmPhotoId + 1, rowCount, 1).getDisplayValues();
  const masterOffset = plates.findIndex((row, offset) =>
    normalizeNopol_(row[0]) === key && normalizeOdometerPhotoId_(photos[offset][0]) === photoId
  );

  if (masterOffset >= 0) {
    const masterRow = master.getRange(masterOffset + 2, 1, 1, master.getLastColumn()).getValues()[0];
    const addedAt = indexes.kmPhotoAddedAt >= 0 ? masterRow[indexes.kmPhotoAddedAt] : '';
    const updatedAt = indexes.kmUpdatedAt >= 0 ? masterRow[indexes.kmUpdatedAt] : '';
    const addedBy = indexes.kmPhotoAddedBy >= 0 ? masterRow[indexes.kmPhotoAddedBy] : '';
    const updatedBy = indexes.kmUpdatedBy >= 0 ? masterRow[indexes.kmUpdatedBy] : '';
    const kmDate = indexes.kmDate >= 0 ? masterRow[indexes.kmDate] : '';
    photoAudit = {
      updatedAt: formatDateTime_(addedAt || updatedAt),
      updatedBy: displayValue_(addedBy || updatedBy),
      kmDate: formatDateOnly_(kmDate),
      source: 'MASTER_UNIT'
    };
  }

  const history = getUnitKmHistorySheet_(ss);
  if ((!photoAudit || !photoAudit.updatedAt || !photoAudit.updatedBy) &&
      history && history.getLastRow() > 1 && historyHeadersReady_(history, UNIT_KM_LOG_HEADERS)) {
    const map = getHistoryHeaderMap_(history, UNIT_KM_LOG_HEADERS);
    const count = history.getLastRow() - 1;
    const historyRows = history.getRange(2, 1, count, history.getLastColumn()).getValues();
    let historyOffset = -1;
    // Pakai evidence terbaru untuk pointer yang sama; history bersifat append-only.
    for (let offset = historyRows.length - 1; offset >= 0; offset--) {
      const row = historyRows[offset];
      if (normalizeNopol_(row[map['NOMOR POLISI'] - 1]) === key && historyOdometerPhotoId_(row, map) === photoId) {
        historyOffset = offset;
        break;
      }
    }
    if (historyOffset >= 0) {
      const historyRow = historyRows[historyOffset];
      const histAddedAt = map['FOTO DITAMBAHKAN AT'] ? historyRow[map['FOTO DITAMBAHKAN AT'] - 1] : '';
      const histTimestamp = map['TIMESTAMP'] ? historyRow[map['TIMESTAMP'] - 1] : '';
      const histAddedBy = map['FOTO DIUNGGAH OLEH'] ? historyRow[map['FOTO DIUNGGAH OLEH'] - 1] : '';
      const histUser = map['USER EMAIL'] ? historyRow[map['USER EMAIL'] - 1] : '';
      const histKmDate = map['TANGGAL KM'] ? historyRow[map['TANGGAL KM'] - 1] : '';
      photoAudit = {
        updatedAt: photoAudit && photoAudit.updatedAt || formatDateTime_(histAddedAt || histTimestamp),
        updatedBy: photoAudit && photoAudit.updatedBy || displayValue_(histAddedBy || histUser),
        kmDate: photoAudit && photoAudit.kmDate || formatDateOnly_(histKmDate),
        source: photoAudit ? 'MASTER_UNIT + ' + UNIT_KM_LOG_SHEET : UNIT_KM_LOG_SHEET
      };
    }
  }

  if (!photoAudit) throw new Error('Foto tidak terkait dengan unit ini.');

  const file = DriveApp.getFileById(photoId);
  const blob = file.getBlob();
  const mimeType = blob.getContentType();
  if (!['image/jpeg','image/png','image/webp'].includes(mimeType)) throw new Error('Format foto tidak didukung.');
  const bytes = blob.getBytes();
  if (!bytes.length || bytes.length > ODOMETER_PHOTO_MAX_BYTES) throw new Error('Ukuran foto tidak valid atau melebihi batas.');
  const base64 = Utilities.base64Encode(bytes);
  if (!base64) throw new Error('Foto gagal dikonversi untuk preview.');
  return {
    mimeType,
    data: base64,
    sizeBytes: bytes.length,
    fileName: file.getName(),
    nopol: normalizeDisplayNopol_(nopol),
    updatedAt: photoAudit.updatedAt || formatDateTime_(file.getDateCreated()),
    updatedBy: photoAudit.updatedBy || '',
    kmDate: photoAudit.kmDate || '',
    source: photoAudit.source || ''
  };
}
/** Jalankan dari editor oleh admin spreadsheet untuk memeriksa akses folder foto. */
function periksaFolderFotoOdometer_() {
  requirePermission_('ADMIN_SETUP', '', { source: 'SHEET' });
  const folder = DriveApp.getFolderById(ODOMETER_PHOTO_FOLDER_ID);
  return { folderId: folder.getId(), folderName: folder.getName(), maxBytes: ODOMETER_PHOTO_MAX_BYTES };
}

/**
 * Histori lama dimigrasikan secara append-only. Urutan dan kolom tambahan lama
 * dipertahankan; seluruh akses berikutnya memakai peta nama header.
 */


// ============================================================================
// 05. SPREADSHEET WORKFLOW & TRIGGERS
// ============================================================================
function ensureUnitKmColumns_(sheet, headers, indexes) {
  const requiredFields = Object.assign({}, UNIT_KM_FIELDS, UNIT_MASTER_DISPLAY_FIELDS);
  const fields = Object.keys(requiredFields).filter(field => indexes[field] < 0);
  if (fields.length) {
    const labels = fields.map(field => requiredFields[field]);
    const requiredColumns = headers.length + labels.length;
    if (requiredColumns > sheet.getMaxColumns()) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
    }
    sheet.getRange(1, headers.length + 1, 1, labels.length).setValues([labels]);
    fields.forEach((field, index) => { indexes[field] = headers.length + index; });
    headers.push(...labels);
  }
  const ordered = ensureMasterUnitWorkflowColumnOrder_(sheet);
  headers.splice(0, headers.length, ...ordered.headers);
  Object.keys(ordered.indexes).forEach(field => { indexes[field] = ordered.indexes[field]; });
}

function buildAliasIndex_(headers, aliasMap) {
  const normalizedHeaders = headers.map(normalizeHeaderKey_);
  const result = {};

  Object.keys(aliasMap).forEach(field => {
    result[field] = -1;

    for (const alias of aliasMap[field]) {
      const index = normalizedHeaders.indexOf(normalizeHeaderKey_(alias));

      if (index >= 0) {
        result[field] = index;
        break;
      }
    }
  });

  return result;
}

function normalizeHeaderKey_(value) {
  return String(value || '')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');
}

function masterUnitValue_(row, index) {
  if (index === undefined || index < 0) {
    return '';
  }

  return displayValue_(row[index]).trim();
}

/**
 * Commit cepat untuk edit langsung Status 3L.
 * Tujuan utamanya adalah menjamin invariant bisnis DONE => Ready to Depart
 * sebelum penulisan ledger, di dalam ScriptLock milik handler. Primary state
 * tetap tersimpan apabila penulisan ledger gagal.
 *
 * Fungsi ini tidak pernah melakukan rollback Status 3L dan tidak mengubah
 * status selain Ready to Depart ketika nilai target benar-benar DONE.
 */
function commitStatus3LPrimaryStateFromSheetEdit_(e, sheet, indexes, actor) {
  if (!e || !e.range || !sheet || !indexes || e.range.getRow() < 2 ||
      indexes.statusUnit < 0 || indexes.status3L < 0 ||
      indexes.status3LUpdatedAt < 0 || indexes.status3LUpdatedBy < 0 ||
      indexes.status3LNote < 0) {
    return { handled: false, rows: [], finalized: [] };
  }
  const startColumn = e.range.getColumn();
  const endColumn = startColumn + e.range.getNumColumns() - 1;
  const status3LColumn = indexes.status3L + 1;
  const status3LNoteColumn = indexes.status3LNote + 1;
  const touchesStatus3L = startColumn <= status3LColumn && endColumn >= status3LColumn;
  const touchesStatus3LNote = startColumn <= status3LNoteColumn && endColumn >= status3LNoteColumn;
  if (!touchesStatus3L && !touchesStatus3LNote) return { handled: false, rows: [], finalized: [] };
  if (!sheetEditStillCurrent_(e)) return { handled: true, rows: [], finalized: [], stale: true };

  const startRow = e.range.getRow();
  const rowCount = e.range.getNumRows();
  const rows = sheet.getRange(startRow, 1, rowCount, sheet.getLastColumn()).getValues();
  const now = new Date();
  const rawUpdater = resolveManualStatus3LEditorEmail_(e, actor);
  const committedRows = [];
  const finalized = [];

  rows.forEach((row, offset) => {
    const rowNumber = startRow + offset;
    const nopol = indexes.nopol >= 0 ? normalizeDisplayNopol_(row[indexes.nopol]) : '';
    if (!nopol) return;
    let status3L = displayValue_(row[indexes.status3L]);
    try {
      status3L = cleanText_(status3L) ? resolveStatus3L_(status3L) : '';
    } catch (error) {
      return;
    }
    const updater = normalizeDoneStatus3LActor_(status3L, rawUpdater);

    // Attribution manual merupakan bagian dari primary state dan tidak menunggu ledger.
    if (touchesStatus3L) {
      sheet.getRange(rowNumber, indexes.status3LUpdatedAt + 1, 1, 2)
        .setValues([[now, safeSheetText_(updater)]]);
    }

    const previousStatusUnit = displayValue_(row[indexes.statusUnit]);
    let nextStatusUnit = previousStatusUnit;
    if (touchesStatus3L && normalizeLookupText_(status3L) === 'DONE') {
      nextStatusUnit = 'Ready to Depart';
      if (normalizeLookupText_(previousStatusUnit) !== normalizeLookupText_('Ready to Depart')) {
        // Installable trigger berjalan dengan otorisasi owner sehingga dapat menulis
        // kolom Status Unit walaupun editor 3L sendiri tidak memiliki izin manual di kolom itu.
        sheet.getRange(rowNumber, indexes.statusUnit + 1).setValue('Ready to Depart');
      }
    }

    committedRows.push({
      rowNumber: rowNumber,
      nopol: nopol,
      calibrationNumber: indexes.calibrationNumber >= 0
        ? parseCalibrationNumber_(row[indexes.calibrationNumber]) : 0,
      status3L: status3L,
      updater: updater,
      previousStatusUnit: previousStatusUnit,
      statusUnit: nextStatusUnit,
      timestamp: now
    });

    if (touchesStatus3L && normalizeLookupText_(status3L) === 'DONE') {
      const finalRow = row.slice();
      finalRow[indexes.status3L] = status3L;
      finalRow[indexes.status3LUpdatedAt] = now;
      finalRow[indexes.status3LUpdatedBy] = updater;
      finalRow[indexes.statusUnit] = nextStatusUnit;
      finalized.push({
        updateId: Utilities.getUuid(), timestamp: now, email: updater, nopol: nopol,
        calibrationNumber: indexes.calibrationNumber >= 0
          ? parseCalibrationNumber_(row[indexes.calibrationNumber]) : 0,
        status3L: status3L, source: 'SPREADSHEET_PRIMARY', reason: 'STATUS_3L_DONE',
        dataAfter: unitKmSnapshot_(finalRow, indexes)
      });
    }
  });
  SpreadsheetApp.flush();

  // FINAL_SNAPSHOT sekarang idempotent. Best-effort di primary phase menjamin
  // kalibrasi DONE tetap terkunci meskipun ledger workflow berikutnya tertunda.
  if (finalized.length) {
    runNonCritical_('FINAL_SNAPSHOT primary Status 3L', () => appendCompletedCalibrationSnapshots_(finalized));
  }
  return { handled: true, rows: committedRows, finalized: finalized };
}

// Event single-cell membawa nilai pada saat edit. Range bisa sudah berisi edit
// orang berikutnya ketika trigger akhirnya berjalan; jangan mengatribusikan nilai
// terbaru itu kepada editor event lama. Bulk event tidak menyediakan values snapshot.
function sheetEditStillCurrent_(e) {
  if (!e || !e.range) return false;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return true;
  const expected = e.value === undefined ? '' : String(e.value);
  return cleanText_(e.range.getValue()) === cleanText_(expected);
}

/**
 * Fallback simple-trigger yang sangat ringan. Ia hanya menjaga state visual
 * DONE => Ready to Depart. Bila Status Unit dilindungi dan simple trigger tidak
 * berhak menulis, kegagalan diabaikan; installable trigger akan melakukan commit
 * yang sama menggunakan otorisasi owner.
 */
function enforceStatus3LDoneReadyInvariantLightweight_(e) {
  try {
    if (!e || !e.range) return 0;
    if (!sheetEditStillCurrent_(e)) return 0;
    const sheet = e.range.getSheet();
    if (!sheet || sheet.getName() !== KM_APP.SHEETS.MASTER_UNIT || e.range.getRow() < 2) return 0;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
    if (indexes.statusUnit < 0 || indexes.status3L < 0) return 0;
    const startColumn = e.range.getColumn();
    const endColumn = startColumn + e.range.getNumColumns() - 1;
    const status3LColumn = indexes.status3L + 1;
    if (!(startColumn <= status3LColumn && endColumn >= status3LColumn)) return 0;
    const startRow = e.range.getRow();
    const rowCount = e.range.getNumRows();
    const statuses3L = sheet.getRange(startRow, indexes.status3L + 1, rowCount, 1).getDisplayValues();
    let changed = 0;
    for (let offset = 0; offset < rowCount; offset++) {
      if (normalizeLookupText_(statuses3L[offset][0]) !== 'DONE') continue;
      try {
        const rowNumber = startRow + offset;
        const live = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
        if (normalizeLookupText_(live[indexes.status3L]) !== 'DONE') continue;
        if (normalizeLookupText_(live[indexes.statusUnit]) !== normalizeLookupText_('Ready to Depart')) {
          sheet.getRange(rowNumber, indexes.statusUnit + 1).setValue('Ready to Depart');
          changed++;
        }
        if (indexes.status3LUpdatedBy >= 0) {
          const currentActor = status3LActorDisplay_(displayValue_(live[indexes.status3LUpdatedBy]));
          if (!isResolvedStatus3LActor_(currentActor) || normalizeLookupText_(currentActor) === 'SYSTEM') {
            sheet.getRange(rowNumber, indexes.status3LUpdatedBy + 1).setValue(KM_STATUS_3L_SHEET_FALLBACK);
            changed++;
          }
        }
      } catch (error) {
        // Simple trigger mungkin tidak memiliki hak menulis kolom terlindungi.
        // Jangan rollback Status 3L; installable trigger tetap menjadi jalur utama.
      }
    }
    if (changed) SpreadsheetApp.flush();
    return changed;
  } catch (error) {
    return 0;
  }
}

function handleMasterUnitSheetWorkflowEdit_(e, sheet, indexes, actor) {
  if (e.range.getRow() < 2) return { handled: false };
  if (!sheetEditStillCurrent_(e)) return { handled: false, stale: true };
  if (indexes.nopol < 0 || indexes.statusUnit < 0 || indexes.calibrationNumber < 0 ||
      indexes.status3L < 0 || indexes.status3LUpdatedAt !== indexes.status3L + 1 ||
      indexes.status3LUpdatedBy !== indexes.status3LUpdatedAt + 1 ||
      indexes.status3LNote !== indexes.status3LUpdatedBy + 1) {
    return { handled: false, databaseNotReady: true };
  }
  const startColumn = e.range.getColumn();
  const endColumn = startColumn + e.range.getNumColumns() - 1;
  const statusUnitColumn = indexes.statusUnit + 1;
  const status3LColumn = indexes.status3L + 1;
  const status3LUpdatedAtColumn = indexes.status3LUpdatedAt + 1;
  const status3LUpdatedByColumn = indexes.status3LUpdatedBy + 1;
  const status3LNoteColumn = indexes.status3LNote + 1;
  const touches = column => column > 0 && startColumn <= column && endColumn >= column;
  const touchesStatusUnit = touches(statusUnitColumn);
  const touchesStatus3L = touches(status3LColumn);
  const touchesStatus3LTimestamp = touches(status3LUpdatedAtColumn);
  const touchesStatus3LUpdatedBy = touches(status3LUpdatedByColumn);
  const touchesStatus3LNote = touches(status3LNoteColumn);
  if (!touchesStatusUnit && !touchesStatus3L && !touchesStatus3LTimestamp &&
      !touchesStatus3LUpdatedBy && !touchesStatus3LNote) {
    return { handled: false };
  }
  if ((touchesStatus3LTimestamp || touchesStatus3LUpdatedBy) && !touchesStatusUnit &&
      !touchesStatus3L && !touchesStatus3LNote) {
    if (e.range.getNumRows() === 1 && e.range.getNumColumns() === 1) {
      e.range.setValue(e.oldValue || '');
    }
    return { handled: true, timestampProtected: true };
  }

  const startRow = e.range.getRow();
  const rowCount = e.range.getNumRows();
  const rows = sheet.getRange(startRow, 1, rowCount, sheet.getLastColumn()).getValues();
  const primaryCommit = e._kmPrimaryStatus3LCommit;
  const now = primaryCommit && primaryCommit.rows.length ? primaryCommit.rows[0].timestamp : new Date();
  const statusEntries = [];
  const status3LEntries = [];
  const completedCalibrationEntries = [];
  const singleEditedCell = rowCount === 1 && e.range.getNumColumns() === 1;
  const latestStatusUnitMap = touchesStatusUnit && !singleEditedCell ? buildLatestStatusUnitMap_() : null;
  const latestStatus3LMap = touchesStatus3L && !singleEditedCell ? buildLatestStatus3LMap_() : null;
  const completedCalibrationKeys = buildCompletedCalibrationKeySet_();

  rows.forEach((row, offset) => {
    const rowNumber = startRow + offset;
    const nopol = normalizeDisplayNopol_(row[indexes.nopol]);
    if (!nopol) return;
    const calibrationNumber = parseCalibrationNumber_(row[indexes.calibrationNumber]);
    const updateId = Utilities.getUuid();
    let statusUnit = displayValue_(row[indexes.statusUnit]);
    let status3L = displayValue_(row[indexes.status3L]);
    let status3LNote = displayValue_(row[indexes.status3LNote]);
    const singleCell = singleEditedCell;
    const primaryRow = e._kmPrimaryStatus3LCommit && e._kmPrimaryStatus3LCommit.rows.find(item => item.rowNumber === rowNumber);
    if (e._kmPrimaryStatus3LCommit && (touchesStatus3L || touchesStatus3LNote) && !primaryRow) return;
    if (primaryRow && (normalizeLookupText_(primaryRow.status3L) !== normalizeLookupText_(status3L) ||
        primaryRow.calibrationNumber !== calibrationNumber)) return;
    const previousStatusUnit = primaryRow ? primaryRow.previousStatusUnit : touchesStatusUnit && singleCell
      ? displayValue_(e.oldValue)
      : (touchesStatusUnit && latestStatusUnitMap && latestStatusUnitMap[normalizeNopol_(nopol)] !== undefined
        ? latestStatusUnitMap[normalizeNopol_(nopol)] : statusUnit);
    let previousStatus3L = touchesStatus3L && singleCell
      ? displayValue_(e.oldValue)
      : (touchesStatus3L && latestStatus3LMap &&
          latestStatus3LMap[normalizeNopol_(nopol) + '|' + calibrationNumber] !== undefined
        ? latestStatus3LMap[normalizeNopol_(nopol) + '|' + calibrationNumber] : status3L);

    if (touchesStatusUnit) {
      const manualStatus = UNIT_MANUAL_OPERATIONAL_STATUSES.find(status =>
        normalizeLookupText_(status) === normalizeLookupText_(statusUnit));
      if (!manualStatus) {
        // v24.1: blank adalah state yang sah untuk unit baru yang belum masuk workflow.
        // Trigger tidak boleh menciptakan Standby hanya karena cell kosong/invalid.
        const previousManualStatus = UNIT_MANUAL_OPERATIONAL_STATUSES.find(status =>
          normalizeLookupText_(status) === normalizeLookupText_(previousStatusUnit)
        ) || '';
        const fallbackStatus = normalizeLookupText_(status3L) === 'DONE'
          ? 'Ready to Depart'
          : previousManualStatus;
        if (displayValue_(statusUnit) !== displayValue_(fallbackStatus)) {
          sheet.getRange(rowNumber, statusUnitColumn).setValue(fallbackStatus);
        }
        return;
      }
      statusUnit = manualStatus;
      const calibrationCompleted = isFinalStatus3L_(status3L) ||
        isCalibrationCompleted_(nopol, calibrationNumber, completedCalibrationKeys);
      const transition3L = calibrationCompleted
        ? { status3L: status3L, changed: false, reason: 'CALIBRATION_DONE_PRESERVED',
            note: 'Siklus kalibrasi sudah selesai; Status 3L tidak diubah oleh perubahan Status Unit.' }
        : nextAutomaticStatus3LForUnit_(status3L, statusUnit, calibrationNumber);
      const nextStatus3L = transition3L.status3L;
      if (transition3L.changed) {
        sheet.getRange(rowNumber, status3LColumn, 1, 4)
          .setValues([[nextStatus3L, now, 'SYSTEM', '']]);
        status3LEntries.push({
          updateId: updateId, timestamp: now, email: 'SYSTEM', nopol: nopol,
          calibrationNumber: calibrationNumber, previousStatus: status3L, status3L: nextStatus3L,
          note: transition3L.note, source: 'AUTO_TRIGGER', reason: transition3L.reason,
          previousStatusUnit: previousStatusUnit, statusUnit: statusUnit
        });
      }
      if (normalizeLookupText_(previousStatusUnit) !== normalizeLookupText_(statusUnit)) {
        statusEntries.push({
          updateId: updateId, timestamp: now, email: actor.email, nopol: nopol, calibrationNumber: calibrationNumber,
          previousStatus: previousStatusUnit, statusUnit: statusUnit,
          note: 'Edit langsung pada MASTER_UNIT.', source: 'SPREADSHEET', reason: 'MANUAL_STATUS_UPDATE'
        });
      }
      return;
    }

    const calibrationCompletedForRow = isFinalStatus3L_(status3L) ||
      isCalibrationCompleted_(nopol, calibrationNumber, completedCalibrationKeys);
    if (normalizeLookupText_(statusUnit) === normalizeLookupText_('On Delivery') && !calibrationCompletedForRow) {
      status3L = '';
      status3LNote = '';
    } else if (touchesStatus3L) {
      try {
        // Explicit blank tetap blank. WAITING hanya boleh dibuat oleh lifecycle
        // kalibrasi yang memang membuka/meresume siklus, bukan sebagai rollback
        // otomatis terhadap edit user.
        status3L = cleanText_(status3L) ? resolveStatus3L_(status3L) : '';
        // Jangan rewrite cell STATUS 3L yang baru saja diedit user. Rewriting
        // nilai yang sama sekalipun akan muncul di Edit history sebagai aksi owner
        // pembuat installable trigger dan membingungkan audit.
      } catch (error) {
        // Data validation Sheet seharusnya menolak status invalid. Trigger tidak
        // pernah menulis previousStatus3L kembali karena itu terlihat sebagai
        // auto-rollback oleh akun owner pembuat trigger.
        runNonCritical_('audit Status 3L invalid tanpa rollback', () => logSystemActivity_(
          nopol, '', 'SHEET_STATUS_3L_INVALID', 'STATUS 3L', previousStatus3L, status3L,
          String(error && error.message ? error.message : error), actor, true));
        return;
      }
    }
    const manualStatus3LMutation = touchesStatus3L || touchesStatus3LNote;
    const primaryStatus3LCommitted = Boolean(e && e._kmPrimaryStatus3LCommit && e._kmPrimaryStatus3LCommit.handled);
    if (touchesStatus3L && !primaryStatus3LCommitted) {
      // Fallback untuk pemanggilan internal/legacy yang tidak melewati primary commit.
      const updater = normalizeDoneStatus3LActor_(status3L, resolveManualStatus3LEditorEmail_(e, actor));
      sheet.getRange(rowNumber, status3LUpdatedAtColumn).setValue(now);
      sheet.getRange(rowNumber, status3LUpdatedByColumn).setValue(safeSheetText_(updater));
    }

    let doneInvariantApplied = false;
    if (primaryRow && normalizeLookupText_(primaryRow.previousStatusUnit) !== normalizeLookupText_(primaryRow.statusUnit)) {
      statusEntries.push({
        updateId: updateId, timestamp: primaryRow.timestamp, email: 'SYSTEM', nopol: nopol,
        calibrationNumber: calibrationNumber, previousStatus: primaryRow.previousStatusUnit,
        statusUnit: primaryRow.statusUnit, note: 'Ready to Depart otomatis karena Status 3L DONE.',
        source: 'AUTO_TRIGGER', reason: 'STATUS_3L_DONE'
      });
    }
    if (manualStatus3LMutation && normalizeLookupText_(status3L) === 'DONE' &&
        normalizeLookupText_(statusUnit) !== normalizeLookupText_('Ready to Depart')) {
      const statusBeforeDone = statusUnit;
      sheet.getRange(rowNumber, statusUnitColumn).setValue('Ready to Depart');
      statusUnit = 'Ready to Depart';
      doneInvariantApplied = true;
      statusEntries.push({
        updateId: updateId, timestamp: now, email: 'SYSTEM', nopol: nopol,
        calibrationNumber: calibrationNumber, previousStatus: statusBeforeDone,
        statusUnit: 'Ready to Depart',
        note: 'Ready to Depart otomatis karena Status 3L DONE.',
        source: 'AUTO_TRIGGER', reason: 'STATUS_3L_DONE'
      });
    }

    if (normalizeLookupText_(statusUnit) !== normalizeLookupText_('On Delivery') && manualStatus3LMutation) {
      const eligibilitySnapshot = unitKmSnapshot_(row, indexes);
      // Validasi harus menggunakan STATUS TARGET yang baru dipilih user. Menggunakan
      // previousStatus3L membuat kasus Ready to Depart + WAITING -> DONE dianggap
      // tidak eligible lalu kolom UPDATED BY tertinggal milik pengaju lama.
      eligibilitySnapshot.status3L = status3L;
      const eligibility = unitStatus3LEligibility_(eligibilitySnapshot);
      if (!eligibility.eligible) {
        const manualUpdater = resolveManualStatus3LEditorEmail_(e, actor);
        // DONE adalah invariant final: jika nilai DONE sudah diterima, Status Unit harus
        // tetap Ready to Depart. Eligibility yang tidak lengkap dicatat sebagai warning,
        // bukan alasan untuk meninggalkan state DONE + Standby yang inkonsisten.
        runNonCritical_('audit Status 3L di luar eligibility tanpa rollback', () => logSystemActivity_(
          nopol, '', 'SHEET_STATUS_3L_NOT_ELIGIBLE', 'STATUS 3L', previousStatus3L, status3L,
          eligibility.reason || 'Status 3L tidak eligible untuk efek workflow otomatis.', actor, true));
        if (normalizeLookupText_(status3L) !== 'DONE') {
          status3LEntries.push({
            updateId: updateId, timestamp: now, email: manualUpdater,
            nopol: nopol, calibrationNumber: calibrationNumber,
            previousStatus: previousStatus3L, status3L: status3L, note: status3LNote,
            source: 'SPREADSHEET', reason: touchesStatus3L ? 'STATUS_3L_UPDATED_NO_WORKFLOW' : 'STATUS_3L_NOTE_UPDATED_NO_WORKFLOW',
            previousStatusUnit: previousStatusUnit, statusUnit: statusUnit
          });
          return;
        }
      }
    }

    // Trigger hanya satu arah: DONE -> Ready to Depart.
    // Koreksi Status 3L tidak pernah mengembalikan Status Unit ke previous state.
    const nextStatusUnit = normalizeLookupText_(status3L) === 'DONE'
      ? 'Ready to Depart'
      : statusUnit;
    const rawStatus3LFinalUpdater = normalizeLookupText_(statusUnit) === normalizeLookupText_('On Delivery') && !calibrationCompletedForRow
      ? 'SYSTEM'
      : resolveManualStatus3LEditorEmail_(e, actor);
    const status3LFinalUpdater = normalizeDoneStatus3LActor_(status3L, rawStatus3LFinalUpdater);
    if (normalizeLookupText_(statusUnit) === normalizeLookupText_('On Delivery') && !calibrationCompletedForRow) {
      sheet.getRange(rowNumber, status3LColumn, 1, 4)
        .setValues([['', now, 'SYSTEM', '']]);
    }
    if (!doneInvariantApplied && normalizeLookupText_(nextStatusUnit) !== normalizeLookupText_(statusUnit)) {
      sheet.getRange(rowNumber, statusUnitColumn).setValue(nextStatusUnit);
      statusEntries.push({
        updateId: updateId, timestamp: now, email: 'SYSTEM', nopol: nopol,
        calibrationNumber: calibrationNumber, previousStatus: statusUnit, statusUnit: nextStatusUnit,
        note: 'Ready to Depart otomatis karena Status 3L DONE.',
        source: 'AUTO_TRIGGER', reason: 'STATUS_3L_DONE'
      });
    }
    status3LEntries.push({
      updateId: updateId, timestamp: primaryRow ? primaryRow.timestamp : now, email: status3LFinalUpdater, nopol: nopol, calibrationNumber: calibrationNumber,
      previousStatus: previousStatus3L, status3L: status3L, note: status3LNote,
      source: 'SPREADSHEET', reason: touchesStatus3L ? 'STATUS_3L_UPDATED' : 'STATUS_3L_NOTE_UPDATED',
      previousStatusUnit: previousStatusUnit, statusUnit: nextStatusUnit
    });
    if (touchesStatus3L && normalizeLookupText_(status3L) === 'DONE' && normalizeLookupText_(previousStatus3L) !== 'DONE') {
      const finalRow = row.slice();
      finalRow[indexes.status3L] = status3L;
      finalRow[indexes.status3LUpdatedAt] = primaryRow ? primaryRow.timestamp : now;
      finalRow[indexes.status3LUpdatedBy] = status3LFinalUpdater;
      finalRow[indexes.status3LNote] = status3LNote;
      finalRow[indexes.statusUnit] = nextStatusUnit;
      completedCalibrationEntries.push({
        updateId: updateId, timestamp: primaryRow ? primaryRow.timestamp : now, email: status3LFinalUpdater, nopol: nopol,
        calibrationNumber: calibrationNumber, status3L: status3L,
        source: 'SPREADSHEET', reason: 'STATUS_3L_DONE',
        dataAfter: unitKmSnapshot_(finalRow, indexes)
      });
    }
  });

  // Audit/history tidak boleh mengubah kembali business state yang sudah dipilih
  // editor. Jika salah satu ledger gagal ditulis, pertahankan cell Status 3L dan
  // catat error di execution log; repair/audit dapat dijalankan sesudahnya.
  runNonCritical_('history Status Unit dari edit Sheet', () => appendStatusUnitHistoryRows_(statusEntries));
  runNonCritical_('history Status 3L dari edit Sheet', () => appendStatus3LHistoryRows_(status3LEntries));
  runNonCritical_('FINAL_SNAPSHOT dari edit Sheet', () => appendCompletedCalibrationSnapshots_(completedCalibrationEntries));
  SpreadsheetApp.flush();
  runNonCritical_('sync analytics dari edit MASTER_UNIT', () => {
    const refreshed = sheet.getRange(startRow, 1, rowCount, sheet.getLastColumn()).getValues();
    updateAnalyticsCalibrationState_(refreshed.map(row => analyticsStateFromMasterRow_(row, indexes)).filter(Boolean));
  });
  return { handled: true, statusUnit: touchesStatusUnit, status3L: touchesStatus3L || touchesStatus3LNote };
}


function handleKalibrasiKmSheetEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  // Edit sheet laporan, history, arsip, atau sheet lain tidak membutuhkan
  // lookup user/RBAC maupun heartbeat trigger. Early return ini mengurangi
  // latensi dan quota pada spreadsheet yang sering diedit.
  const managedSheets = [
    KM_APP.SHEETS.USERS,
    KM_APP.SHEETS.ROLES,
    KM_APP.SHEETS.MASTER_UNIT,
    KM_APP.SHEETS.CONFIG
  ];
  if (!managedSheets.includes(sheetName)) return;
  // Only the canonical installed trigger processes the event. Other accounts'
  // triggers are invisible to getProjectTriggers(), but share this project property.
  if (e.triggerUid) {
    const canonicalId = PropertiesService.getScriptProperties().getProperty(KM_SHEET_EDIT_TRIGGER.ID_PROPERTY);
    if (canonicalId && String(e.triggerUid) !== canonicalId) return;
  }
  try {
    PropertiesService.getScriptProperties().setProperty(
      KM_SHEET_EDIT_TRIGGER.LAST_FIRED_PROPERTY, new Date().toISOString()
    );
  } catch (heartbeatError) {}
  let actor = getCurrentUserContext_({ allowUnauthorized: true, source: 'SHEET', event: e, forceRevalidate: true });

  if (sheetName === KM_APP.SHEETS.USERS) {
    invalidateUserDirectoryCache_();
    runNonCritical_('audit edit KM_USERS', () => logSystemActivity_('', '',
      actor.permissions.includes('MANAGE_USERS') ? 'SHEET_USER_EDIT' : 'UNAUTHORIZED_SHEET_EDIT',
      'KM_USERS', e.oldValue || '', e.value || '', 'Edit langsung pada sheet pengguna.', actor));
    return;
  }

  if (sheetName === KM_APP.SHEETS.ROLES) {
    invalidateSecurityCaches_();
    runNonCritical_('audit edit KM_ROLES', () => logSystemActivity_('', '',
      actor.permissions.includes('MANAGE_USERS') ? 'SHEET_ROLE_EDIT' : 'UNAUTHORIZED_SHEET_EDIT',
      'KM_ROLES', e.oldValue || '', e.value || '', 'Edit langsung pada sheet role.', actor));
    return;
  }

  if (sheetName === KM_APP.SHEETS.MASTER_UNIT) {
    invalidateAppCaches_();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
    const status3LColumn = indexes.status3L >= 0 ? indexes.status3L + 1 : -1;
    const status3LUpdatedAtColumn = indexes.status3LUpdatedAt >= 0 ? indexes.status3LUpdatedAt + 1 : -1;
    const status3LUpdatedByColumn = indexes.status3LUpdatedBy >= 0 ? indexes.status3LUpdatedBy + 1 : -1;
    const status3LNoteColumn = indexes.status3LNote >= 0 ? indexes.status3LNote + 1 : -1;
    const editStartColumn = e.range.getColumn();
    const editEndColumn = editStartColumn + e.range.getNumColumns() - 1;
    const onlyStatus3LFields = e.range.getRow() > 1 && status3LColumn > 0 &&
      status3LUpdatedAtColumn === status3LColumn + 1 &&
      status3LUpdatedByColumn === status3LUpdatedAtColumn + 1 &&
      status3LNoteColumn === status3LUpdatedByColumn + 1 &&
      editStartColumn >= status3LColumn && editEndColumn <= status3LNoteColumn;
    const workflowReady = indexes.nopol >= 0 && indexes.statusUnit >= 0 &&
      indexes.calibrationNumber >= 0 && status3LUpdatedAtColumn === status3LColumn + 1 &&
      status3LUpdatedByColumn === status3LUpdatedAtColumn + 1 &&
      status3LNoteColumn === status3LUpdatedByColumn + 1;
    const touchesStatus3L = onlyStatus3LFields &&
      editStartColumn <= status3LColumn && editEndColumn >= status3LColumn;
    const touchesStatus3LNote = onlyStatus3LFields && status3LNoteColumn > 0 &&
      editStartColumn <= status3LNoteColumn && editEndColumn >= status3LNoteColumn;
    let masterEditAuthorized = onlyStatus3LFields
      ? actor.permissions.includes('UPDATE_STATUS_3L')
      : actor.permissions.includes('UPDATE_UNIT');

    if (masterEditAuthorized && !actor.globalScope && e.range.getRow() > 1) {
      masterEditAuthorized = indexes.operatingPoint >= 0;
      const start = e.range.getRow();
      const count = e.range.getNumRows();
      const opValues = masterEditAuthorized
        ? sheet.getRange(start, indexes.operatingPoint + 1, count, 1).getDisplayValues() : [];
      for (let offset = 0; masterEditAuthorized && offset < count; offset++) {
        masterEditAuthorized = canAccessOperatingPoint_(actor, opValues[offset][0]);
      }
    }

    if (!masterEditAuthorized) {
      // Proteksi range sengaja tidak digunakan. Untuk edit satu sel, trigger
      // memulihkan nilai lama bila actor tidak lolos RBAC/scope validation.
      if (e.range.getNumRows() === 1 && e.range.getNumColumns() === 1 &&
          e.oldValue !== undefined) e.range.setValue(e.oldValue || '');
      runNonCritical_('audit edit MASTER_UNIT ditolak', () => logSystemActivity_('', '',
        'UNAUTHORIZED_SHEET_EDIT', onlyStatus3LFields ? 'STATUS 3L / CATATAN' : 'MASTER_UNIT',
        e.oldValue || '', e.value || '',
        'Edit tidak diproses workflow karena actor tidak lolos RBAC/scope; nilai lama dipulihkan bila tersedia.', actor));
      return;
    }

    if (workflowReady) {
      let workflowLock = null;
      try {
        workflowLock = acquireScriptLock_(30000, 'sinkronisasi edit MASTER_UNIT');
        if (!sheetEditStillCurrent_(e)) return;
        if (onlyStatus3LFields) {
          e._kmPrimaryStatus3LCommit = commitStatus3LPrimaryStateFromSheetEdit_(e, sheet, indexes, actor);
        }
        handleMasterUnitSheetWorkflowEdit_(e, sheet, indexes, actor);
        if (e._kmPrimaryStatus3LCommit && e._kmPrimaryStatus3LCommit.rows.length) {
          // Evidence per unit tetap berguna jika penulisan STATUS_3L_HISTORY gagal.
          const commit = e._kmPrimaryStatus3LCommit;
          runNonCritical_('activity per unit Status 3L', () => logSystemActivityBatch_(commit.rows.map(item => ({
            nopol: item.nopol, action: touchesStatus3L ? 'SHEET_STATUS_3L_EDIT' : 'SHEET_STATUS_3L_NOTE_EDIT',
            fieldName: touchesStatus3L ? 'STATUS 3L' : 'CATATAN STATUS 3L',
            oldValue: e.oldValue || '', newValue: item.status3L,
            actor: Object.assign({}, actor, { email: item.updater }),
            details: { calibrationNumber: item.calibrationNumber, rowNumber: item.rowNumber,
              editedAt: item.timestamp.toISOString(), triggerUid: cleanText_(e.triggerUid) }
          })), actor, true));
        }
      } catch (workflowError) {
        enforceStatus3LDoneReadyInvariantLightweight_(e);
        // Jangan rollback edit user. Jika lock gagal sebelum primary commit,
        // execution log ini menjadi diagnosis untuk retry/repair workflow.
        console.error('Workflow history MASTER_UNIT tertunda: ' +
          String(workflowError && workflowError.message ? workflowError.message : workflowError));
      } finally {
        if (workflowLock) workflowLock.releaseLock();
        invalidateAppCaches_();
      }
    }

    runNonCritical_('audit edit MASTER_UNIT', () => logSystemActivity_('', '',
      workflowReady ? (onlyStatus3LFields ? 'SHEET_STATUS_3L_EDIT' : 'SHEET_MASTER_UNIT_EDIT') : 'SHEET_DATABASE_NOT_READY',
      onlyStatus3LFields ? 'STATUS 3L / CATATAN' : 'MASTER_UNIT', e.oldValue || '', e.value || '', 'Edit langsung pada master unit.', actor));
    return;
  }

  if (sheetName === KM_APP.SHEETS.CONFIG) {
    invalidateAppCaches_();
    runNonCritical_('refresh validasi Status 3L', () => {
      const master = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
      if (!master || !master.getLastColumn()) return;
      const headers = master.getRange(1, 1, 1, master.getLastColumn()).getDisplayValues()[0];
      applyStatus3LValidation_(master, buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES));
    });
    runNonCritical_('audit edit MASTER_CONFIG', () => logSystemActivity_('', '',
      actor.permissions.includes('ADMIN_SETUP') ? 'SHEET_CONFIG_EDIT' : 'UNAUTHORIZED_SHEET_EDIT',
      'MASTER_CONFIG', e.oldValue || '', e.value || '', 'Edit langsung pada konfigurasi.', actor));
    return;
  }


}

/**
 * Fallback untuk project yang belum sempat menjalankan migrasi terbaru.
 * Setelah trigger installable terpasang, wrapper simple trigger tidak ikut
 * memproses event agar history tidak tercatat dua kali.
 */
function onEdit(e) {
  let installed = false;
  try {
    installed = PropertiesService.getScriptProperties()
      .getProperty(KM_SHEET_EDIT_TRIGGER.ENABLED_PROPERTY) === 'TRUE';
  } catch (error) {}
  // Satu jalur penulis: simple trigger tidak boleh mendahului installable trigger
  // dan menghilangkan STATUS UNIT SEBELUM dari history.
  if (!installed) handleKalibrasiKmSheetEdit(e);
}

function auditStatus3LAutoRollbackRisk() {
  const actor = requirePermission_('ADMIN_SETUP', '', { source: 'SHEET' });
  KM_REQUEST_ACTOR_ = actor;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  const triggers = ScriptApp.getProjectTriggers().filter(trigger =>
    trigger.getHandlerFunction() === KM_SHEET_EDIT_TRIGGER.HANDLER
  );
  const users = readUserDirectory_(true).filter(user => user.active).map(user => ({
    email: user.email,
    roles: user.roles,
    operatingPoints: user.operatingPoints,
    canUpdateStatus3L: permissionsForRoles_(user.roles).includes('UPDATE_STATUS_3L')
  }));
  let protections = [];
  if (master) {
    protections = master.getProtections(SpreadsheetApp.ProtectionType.RANGE)
      .filter(item => String(item.getDescription() || '').indexOf('KM_RBAC:STATUS_3L') === 0)
      .map(item => ({
        description: item.getDescription(),
        range: item.getRange().getA1Notation(),
        warningOnly: item.isWarningOnly(),
        editors: item.getEditors().map(editor => normalizeEmail_(editor.getEmail())).filter(Boolean)
      }));
  }
  const regression = auditStatus3LDoneRegression();
  return {
    success: true,
    ownerEmail: getSpreadsheetOwnerEmail_(),
    triggerCount: triggers.length,
    triggerHandlers: triggers.map(trigger => trigger.getHandlerFunction()),
    status3LProtections: protections,
    users: users,
    doneRegressionCount: regression.count,
    doneRegressions: regression.rows
  };
}

function applyMasterUnitStatusValidation_(sheet, statusUnitIndex) {
  if (!sheet || statusUnitIndex < 0 || sheet.getMaxRows() < 2) return false;
  const range = sheet.getRange(2, statusUnitIndex + 1, Math.max(sheet.getMaxRows() - 1, 1), 1);
  range.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(UNIT_OPERATIONAL_STATUSES, true)
    .setAllowInvalid(false)
    .setHelpText('Standby / On Delivery dapat dipilih manual. Ready to Depart otomatis saat Status 3L DONE.')
    .build());
  return true;
}

function auditStatus3LWorkflowIntegrity_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  const report = {
    sheetReady: false,
    statusValidationReady: false,
    triggerCount: 0,
    externalWatchTriggerCount: 0,
    externalWatchReady: false,
    initialBlankUnits: [],
    preCalibrationStatusCandidates: [],
    doneWithoutReadyToDepart: [],
    doneAttributionNotTeam3L: [],
    unresolvedManualAttribution: [],
    systemManualAttribution: []
  };
  try {
    report.triggerCount = ScriptApp.getProjectTriggers().filter(trigger =>
      trigger.getHandlerFunction() === KM_SHEET_EDIT_TRIGGER.HANDLER
    ).length;
  } catch (error) {}
  try {
    report.externalWatchTriggerCount = ScriptApp.getProjectTriggers().filter(trigger =>
      trigger.getHandlerFunction() === KM_STATUS_3L_WATCH.HANDLER &&
      trigger.getEventType() === ScriptApp.EventType.CLOCK
    ).length;
    report.externalWatchReady = isStatus3LWatchTriggerReady_();
  } catch (error) {}
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return report;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  report.sheetReady = indexes.nopol >= 0 && indexes.statusUnit >= 0 && indexes.status3L >= 0 &&
    indexes.status3LUpdatedBy >= 0 && indexes.calibrationNumber >= 0;
  if (!report.sheetReady) return report;
  report.statusValidationReady = isMasterUnitStatusValidationReady_(sheet, indexes.statusUnit);
  if (sheet.getLastRow() < 2) return report;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  rows.forEach((row, offset) => {
    const nopol = normalizeDisplayNopol_(row[indexes.nopol]);
    if (!nopol) return;
    const status3L = displayValue_(row[indexes.status3L]);
    const statusUnit = displayValue_(row[indexes.statusUnit]);
    const rawUpdatedBy = displayValue_(row[indexes.status3LUpdatedBy]);
    const updatedBy = status3LActorDisplay_(rawUpdatedBy);
    const calibrationNumber = parseCalibrationNumber_(row[indexes.calibrationNumber]);
    const item = { row: offset + 2, nopol: nopol,
      calibrationNumber: calibrationNumber,
      statusUnit: statusUnit, status3L: status3L, updatedBy: updatedBy };

    // State awal valid: unit sudah ada di master tetapi belum pernah masuk workflow.
    if (!calibrationNumber && !cleanText_(status3L)) {
      if (!cleanText_(statusUnit)) {
        if (report.initialBlankUnits.length < 500) report.initialBlankUnits.push(item);
      } else if (report.preCalibrationStatusCandidates.length < 500) {
        // Hanya kandidat review. Jangan auto-clear karena status bisa merupakan edit manual yang sah.
        report.preCalibrationStatusCandidates.push(item);
      }
    }

    if (normalizeLookupText_(status3L) === 'DONE' &&
        normalizeLookupText_(statusUnit) !== normalizeLookupText_('Ready to Depart')) {
      if (report.doneWithoutReadyToDepart.length < 500) report.doneWithoutReadyToDepart.push(item);
    }
    if (normalizeLookupText_(status3L) === 'DONE' &&
        (!isResolvedStatus3LActor_(updatedBy) || normalizeLookupText_(updatedBy) === 'SYSTEM')) {
      if (report.doneAttributionNotTeam3L.length < 500) report.doneAttributionNotTeam3L.push(item);
    }
    if (normalizeLookupText_(rawUpdatedBy) === normalizeLookupText_(KM_STATUS_3L_LEGACY_UNRESOLVED)) {
      if (report.unresolvedManualAttribution.length < 500) report.unresolvedManualAttribution.push(item);
    }
    if (normalizeLookupText_(updatedBy) === 'SYSTEM' && normalizeLookupText_(status3L) === 'DONE') {
      if (report.systemManualAttribution.length < 500) report.systemManualAttribution.push(item);
    }
  });
  return report;
}

function normalizeCompletedHistoryDoneAttribution_() {
  const sheet = getUnitKmHistorySheet_(SpreadsheetApp.getActiveSpreadsheet());
  if (!sheet || sheet.getLastRow() < 2 || !historyHeadersReady_(sheet, UNIT_KM_LOG_HEADERS)) {
    return { changed: 0, rows: 0 };
  }
  const map = getHistoryHeaderMap_(sheet, UNIT_KM_LOG_HEADERS);
  const rowCount = sheet.getLastRow() - 1;
  const rows = sheet.getRange(2, 1, rowCount, sheet.getLastColumn()).getValues();
  const updatedByValues = sheet.getRange(2, map['STATUS 3L UPDATED BY'], rowCount, 1).getValues();
  const completedByValues = sheet.getRange(2, map['COMPLETED BY'], rowCount, 1).getValues();
  const dataAfterValues = sheet.getRange(2, map['DATA SETELAH'], rowCount, 1).getValues();
  let changed = 0;

  rows.forEach((row, offset) => {
    const type = normalizeLookupText_(getRowValue_(row, map, 'RECORD TYPE'));
    const lifecycle = normalizeLookupText_(getRowValue_(row, map, 'LIFECYCLE STATUS'));
    if (type !== 'FINAL_SNAPSHOT' || lifecycle !== 'LOCKED') return;

    const rawUpdated = status3LActorDisplay_(updatedByValues[offset][0]);
    const rawCompleted = status3LActorDisplay_(completedByValues[offset][0]);
    const rowActor = status3LActorDisplay_(getRowValue_(row, map, 'USER EMAIL'));
    const resolved = preferredDoneStatus3LActor_(
      rawUpdated, preferredDoneStatus3LActor_(rawCompleted, rowActor)
    );

    let rowChanged = false;
    if (normalizeLookupText_(rawUpdated) !== normalizeLookupText_(resolved)) {
      updatedByValues[offset][0] = resolved;
      rowChanged = true;
    }
    if (normalizeLookupText_(rawCompleted) !== normalizeLookupText_(resolved)) {
      completedByValues[offset][0] = resolved;
      rowChanged = true;
    }

    const rawDataAfter = cleanText_(dataAfterValues[offset][0]);
    if (rawDataAfter) {
      try {
        const parsed = JSON.parse(rawDataAfter);
        if (parsed && typeof parsed === 'object') {
          const dataActor = status3LActorDisplay_(parsed.status3LUpdatedBy);
          if (normalizeLookupText_(dataActor) !== normalizeLookupText_(resolved) ||
              normalizeLookupText_(parsed.status3L) !== 'DONE' ||
              normalizeLookupText_(parsed.statusUnit) !== normalizeLookupText_('Ready to Depart')) {
            parsed.status3LUpdatedBy = resolved;
            parsed.status3L = 'DONE';
            parsed.statusUnit = 'Ready to Depart';
            dataAfterValues[offset][0] = JSON.stringify(parsed);
            rowChanged = true;
          }
        }
      } catch (error) {}
    }
    if (rowChanged) changed++;
  });

  if (changed) {
    sheet.getRange(2, map['STATUS 3L UPDATED BY'], rowCount, 1).setValues(updatedByValues);
    sheet.getRange(2, map['COMPLETED BY'], rowCount, 1).setValues(completedByValues);
    sheet.getRange(2, map['DATA SETELAH'], rowCount, 1).setValues(dataAfterValues);
  }
  return { changed: changed, rows: rowCount };
}

/**
 * Repair khusus invariant DONE => Ready to Depart dan attribution DONE yang valid.
 * Email editor dipertahankan bila tersedia; fallback Team 3L hanya dipakai bila
 * identitas editor memang tidak dapat di-resolve. Sumber audit teknis tetap tersimpan.
 */
function repairStatus3LDoneReadyInvariant_(actor, options) {
  const settings = options || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  if (!sheet || sheet.getLastRow() < 2 || !sheet.getLastColumn()) {
    return { fixed: 0, finalSnapshots: 0, unsupportedDone: 0, normalizedAttribution: 0, historyAttributionNormalized: 0, reconciliation: { changed: 0 }, rows: [] };
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  if (indexes.nopol < 0 || indexes.statusUnit < 0 || indexes.status3L < 0 || indexes.calibrationNumber < 0) {
    throw new Error('MASTER_UNIT belum memiliki kolom workflow lengkap.');
  }
  applyMasterUnitStatusValidation_(sheet, indexes.statusUnit);

  // Preserve primary status, enforce DONE invariants, and report missing history
  // separately. A delayed ledger must never downgrade a user's current status.
  const reconciliation = settings.skipReconcile
    ? { changed: 0, skippedAlreadyReconciled: true }
    : reconcileMasterUnitWorkflowState_(sheet, indexes);
  const trustedStatus = buildTrustedStatus3LStateIndex_(ss.getSheetByName(UNIT_STATUS_3L_LOG_SHEET));
  const trustedFinals = buildTrustedFinalSnapshotIndex_(trustedStatus);
  const completedKeys = buildCompletedCalibrationKeySet_();
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const fixes = [];
  const attributionFixes = [];
  const unsupported = [];
  const finalEntries = [];

  rows.forEach((row, offset) => {
    const nopol = normalizeDisplayNopol_(row[indexes.nopol]);
    const sequence = parseCalibrationNumber_(row[indexes.calibrationNumber]);
    if (!nopol || normalizeLookupText_(row[indexes.status3L]) !== 'DONE') return;
    const key = calibrationCompletionKey_(nopol, sequence);
    const bucket = trustedStatus[key] || null;
    const latest = bucket && bucket.latest || null;
    const trustedCurrentDone = latest && normalizeLookupText_(latest.status3L) === 'DONE';
    if (!trustedCurrentDone && !completedKeys.has(key)) {
      unsupported.push({ row: offset + 2, nopol: nopol, calibrationNumber: sequence });
    }
    const needsUnit = normalizeLookupText_(row[indexes.statusUnit]) !== normalizeLookupText_('Ready to Depart');
    const currentDoneActor = indexes.status3LUpdatedBy >= 0
      ? status3LActorDisplay_(displayValue_(row[indexes.status3LUpdatedBy])) : '';
    const targetDoneActor = preferredDoneStatus3LActor_(
      currentDoneActor, bucket && bucket.done ? bucket.done.updatedBy : ''
    );
    const needsAttribution = indexes.status3LUpdatedBy >= 0 &&
      normalizeLookupText_(currentDoneActor) !== normalizeLookupText_(targetDoneActor);
    const needsSnapshot = Boolean(!trustedFinals[key] && bucket && bucket.done);
    // Avoid thousands of individual reads for already-consistent completed rows.
    if (!needsUnit && !needsAttribution && !needsSnapshot) return;
    const live = sheet.getRange(offset + 2, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!masterWorkflowRowUnchanged_(row, live, indexes)) return;
    const previousStatusUnit = displayValue_(row[indexes.statusUnit]);
    if (normalizeLookupText_(previousStatusUnit) !== normalizeLookupText_('Ready to Depart')) {
      sheet.getRange(offset + 2, indexes.statusUnit + 1).setValue('Ready to Depart');
      fixes.push({ rowNumber: offset + 2, nopol: nopol, calibrationNumber: sequence, previousStatusUnit: previousStatusUnit });
    }
    if (needsAttribution) {
      sheet.getRange(offset + 2, indexes.status3LUpdatedBy + 1).setValue(safeSheetText_(targetDoneActor));
      attributionFixes.push({ rowNumber: offset + 2, nopol: nopol, calibrationNumber: sequence,
        previousUpdatedBy: displayValue_(row[indexes.status3LUpdatedBy]),
        updatedBy: targetDoneActor, statusUnit: 'Ready to Depart' });
    }
    if (needsSnapshot) {
      const finalRow = row.slice();
      finalRow[indexes.statusUnit] = 'Ready to Depart';
      finalRow[indexes.status3LUpdatedBy] = targetDoneActor;
      finalEntries.push({
        updateId: Utilities.getUuid(),
        timestamp: bucket.done.timestamp || new Date(),
        email: bucket.done.updatedBy || KM_STATUS_3L_SHEET_FALLBACK,
        nopol: nopol,
        calibrationNumber: sequence,
        status3L: 'DONE',
        source: 'SYSTEM_REPAIR',
        reason: 'REPAIR_FINAL_SNAPSHOT_FROM_TRUSTED_DONE',
        dataAfter: unitKmSnapshot_(finalRow, indexes)
      });
    }
  });

  if (fixes.length) appendStatusUnitHistoryRows_(fixes.map(item => ({
    updateId: Utilities.getUuid(), timestamp: new Date(), email: 'SYSTEM', nopol: item.nopol,
    calibrationNumber: item.calibrationNumber, previousStatus: item.previousStatusUnit,
    statusUnit: 'Ready to Depart', note: 'Repair invariant dari completion evidence tepercaya: Status 3L DONE => Ready to Depart.',
    source: 'SYSTEM_REPAIR', reason: 'STATUS_3L_DONE_INVARIANT_REPAIR'
  })));
  if (attributionFixes.length) appendStatus3LHistoryRows_(attributionFixes.map(item => ({
    updateId: Utilities.getUuid(), timestamp: new Date(), email: item.updatedBy || KM_STATUS_3L_SHEET_FALLBACK,
    nopol: item.nopol, calibrationNumber: item.calibrationNumber,
    previousStatus: 'DONE', status3L: 'DONE',
    note: 'Attribution Status 3L DONE diselaraskan tanpa menghapus email editor yang valid.',
    source: 'ADMIN_NORMALIZATION', reason: 'STATUS_3L_DONE_ATTRIBUTION_ALIGNED',
    previousStatusUnit: item.statusUnit, statusUnit: 'Ready to Depart'
  })));
  const historyAttribution = normalizeCompletedHistoryDoneAttribution_();
  const finalResult = finalEntries.length ? appendCompletedCalibrationSnapshots_(finalEntries) : null;
  SpreadsheetApp.flush();
  invalidateAppCaches_();
  return {
    fixed: fixes.length,
    finalSnapshots: finalResult && Number(finalResult.count) ? Number(finalResult.count) : 0,
    skippedExistingSnapshots: finalResult && Number(finalResult.skippedExisting) ? Number(finalResult.skippedExisting) : 0,
    unsupportedDone: unsupported.length,
    unsupportedRows: unsupported.slice(0, 100),
    normalizedAttribution: attributionFixes.length,
    historyAttributionNormalized: historyAttribution.changed,
    reconciliation: reconciliation,
    rows: fixes.slice(0, 500)
  };
}

function repairStatus3LDoneReadyInvariant() {
  const actor = requirePermission_('ADMIN_SETUP', '', { source: 'SHEET' });
  const lock = acquireScriptLock_(30000, 'repair invariant Status 3L DONE');
  try { return repairStatus3LDoneReadyInvariant_(actor); }
  finally { lock.releaseLock(); }
}

function repairStatus3LWorkflowDataIntegrityUnsafe_(actor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  if (!sheet || !sheet.getLastColumn()) throw new Error('MASTER_UNIT belum tersedia.');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  if (indexes.statusUnit < 0 || indexes.status3L < 0 || indexes.nopol < 0) {
    throw new Error('MASTER_UNIT belum memiliki kolom workflow yang lengkap.');
  }

  applyMasterUnitStatusValidation_(sheet, indexes.statusUnit);
  applyStatus3LValidation_(sheet, indexes);
  const trigger = ensureKalibrasiKmEditTrigger_();
  const status3LWatchTrigger = ensureStatus3LWatchTrigger_();
  invalidateSecurityCaches_();
  KM_REQUEST_ROLE_DIRECTORY_ = null;
  KM_REQUEST_USER_DIRECTORY_ = null;
  KM_REQUEST_ACTIVE_ROLE_NAMES_ = null;
  const accessMode = removeAutomaticSheetProtections_();

  KM_REQUEST_ACTOR_ = actor || KM_REQUEST_ACTOR_;
  const reconciled = reconcileMasterUnitWorkflowState_(sheet, indexes);
  const status3LLog = ss.getSheetByName(UNIT_STATUS_3L_LOG_SHEET);
  const attribution = status3LLog
    ? backfillStatus3LTimestamps_(sheet, indexes, status3LLog)
    : { changed: 0, matched: 0, unresolved: 0, mismatches: [] };
  const doneReadyInvariant = repairStatus3LDoneReadyInvariant_(actor, { skipReconcile: true });

  SpreadsheetApp.flush();
  invalidateAppCaches_();
  return {
    validationApplied: true,
    trigger: trigger,
    status3LWatchTrigger: status3LWatchTrigger,
    accessMode: accessMode,
    doneReadyInvariant: doneReadyInvariant,
    reconciled: reconciled,
    attribution: attribution,
    audit: auditStatus3LWorkflowIntegrity_()
  };
}

// Alias compatibility v3: nama lama sekarang safe trigger-only.
// Repair data workflow historis dipisahkan ke repairStatus3LWorkflowDataIntegrityUnsafe_
// dan tidak dipanggil oleh tombol trigger Administration.
function repairStatus3LWorkflowIntegrity_(actor) {
  return repairKalibrasiKmTriggersOnly_();
}

function repairKalibrasiKmTriggersOnly_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  let validation = { statusUnit: false, status3L: false };
  if (sheet && sheet.getLastColumn()) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
    if (indexes.statusUnit >= 0) validation.statusUnit = applyMasterUnitStatusValidation_(sheet, indexes.statusUnit);
    if (indexes.status3L >= 0) {
      applyStatus3LValidation_(sheet, indexes);
      validation.status3L = true;
    }
  }
  const editTrigger = ensureKalibrasiKmEditTrigger_();
  const status3LWatchTrigger = ensureStatus3LWatchTrigger_();
  SpreadsheetApp.flush();
  invalidateAppCaches_();
  return {
    success: true,
    safeTriggerOnly: true,
    existingDataChanged: false,
    editTrigger: editTrigger,
    status3LWatchTrigger: status3LWatchTrigger,
    validation: validation,
    message: 'Trigger + validation disinkronkan. Nilai/status/history existing tidak direkonsiliasi atau di-repair.'
  };
}

function repairStatus3LWorkflowIntegrity() {
  requirePermission_('ADMIN_SETUP', '', { source: 'SHEET' });
  const lock = acquireScriptLock_(30000, 'sinkronisasi trigger Status 3L aman');
  try { return repairKalibrasiKmTriggersOnly_(); }
  finally { lock.releaseLock(); }
}

function repairKalibrasiKmEditTrigger() {
  requirePermission_('ADMIN_SETUP', '', { source: 'SHEET' });
  const lock = acquireScriptLock_(30000, 'sinkronisasi trigger Kalibrasi KM aman');
  try { return repairKalibrasiKmTriggersOnly_(); }
  finally { lock.releaseLock(); }
}


/**
 * Audit read-only khusus rule v24.1:
 * - unit baru boleh STATUS UNIT kosong;
 * - canonical installable trigger harus tunggal;
 * - kandidat status pre-calibration ditampilkan untuk review tanpa auto-clear.
 */
function auditInitialUnitStatusIntegrity() {
  requirePermission_('ADMIN_SETUP', '', { source: 'SHEET' });
  const report = auditStatus3LWorkflowIntegrity_();
  return {
    success: true,
    triggerReady: report.triggerCount === 1,
    triggerCount: report.triggerCount,
    statusValidationReady: report.statusValidationReady,
    initialBlankCount: report.initialBlankUnits.length,
    preCalibrationStatusCandidateCount: report.preCalibrationStatusCandidates.length,
    initialBlankUnits: report.initialBlankUnits,
    preCalibrationStatusCandidates: report.preCalibrationStatusCandidates,
    doneWithoutReadyToDepart: report.doneWithoutReadyToDepart
  };
}

function repairStatus3LNonOwnerWorkflow() {
  requirePermission_('ADMIN_SETUP', '', { source: 'SHEET' });
  const lock = acquireScriptLock_(30000, 'sinkronisasi trigger non-owner aman');
  try {
    const result = repairKalibrasiKmTriggersOnly_();
    result.legacyFunction = true;
    result.message = 'Fungsi legacy dialihkan ke safe trigger-only mode; data existing tidak diubah.';
    return result;
  } finally { lock.releaseLock(); }
}


/**
 * Watcher untuk perubahan Status 3L yang berasal dari automation / Google Sheets API.
 * Perubahan programmatic tidak memicu onEdit, sehingga primary state dapat berubah
 * sementara STATUS_3L_HISTORY dan kolom UPDATED AT/BY tetap tertinggal.
 *
 * Strategy:
 * - polling ringan setiap 1 menit hanya terhadap projection MASTER_UNIT;
 * - full history hanya dibaca jika signature business-state berubah;
 * - beri grace 2 menit agar onEdit manual sempat menulis ledger;
 * - jika divergence tetap ada, catat sebagai EXTERNAL_SYNC dan update timestamp.
 */
function status3LWatchBusinessProjection_(sheet, indexes) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const required = [indexes.nopol, indexes.statusUnit, indexes.calibrationNumber, indexes.status3L];
  if (required.some(index => index < 0)) return [];
  const rowCount = sheet.getLastRow() - 1;

  // Hot path setiap menit: hanya 4 kolom bisnis dibaca (bukan seluruh B:R).
  // Ini menjaga watcher ringan walaupun MASTER_UNIT memiliki ribuan unit.
  const nopols = sheet.getRange(2, indexes.nopol + 1, rowCount, 1).getDisplayValues();
  const statusUnits = sheet.getRange(2, indexes.statusUnit + 1, rowCount, 1).getDisplayValues();
  const workflowWidth = indexes.status3L - indexes.calibrationNumber + 1;
  const workflow = workflowWidth > 0
    ? sheet.getRange(2, indexes.calibrationNumber + 1, rowCount, workflowWidth).getDisplayValues()
    : [];
  const calibrationOffset = 0;
  const status3LOffset = indexes.status3L - indexes.calibrationNumber;

  return nopols.map((row, offset) => ({
    rowNumber: offset + 2,
    nopol: normalizeDisplayNopol_(row[0]),
    calibrationNumber: parseCalibrationNumber_(workflow[offset] && workflow[offset][calibrationOffset]),
    statusUnit: displayValue_(statusUnits[offset] && statusUnits[offset][0]),
    status3L: displayValue_(workflow[offset] && workflow[offset][status3LOffset])
  })).filter(item => item.nopol);
}

function status3LWatchAttachMetadata_(sheet, indexes, projection) {
  const rows = Array.isArray(projection) ? projection : [];
  if (!rows.length || indexes.status3LUpdatedAt < 0 || indexes.status3LUpdatedBy < 0 ||
      indexes.status3LNote < 0) return rows;
  const rowCount = sheet.getLastRow() - 1;
  const width = indexes.status3LNote - indexes.status3LUpdatedAt + 1;
  const metadata = sheet.getRange(2, indexes.status3LUpdatedAt + 1, rowCount, width).getValues();
  rows.forEach(item => {
    const source = metadata[item.rowNumber - 2] || [];
    item.status3LUpdatedAt = source[0] || '';
    item.status3LUpdatedBy = status3LActorDisplay_(displayValue_(source[indexes.status3LUpdatedBy - indexes.status3LUpdatedAt]));
    item.status3LNote = displayValue_(source[indexes.status3LNote - indexes.status3LUpdatedAt]);
  });
  return rows;
}

function status3LWatchSignature_(projection) {
  const state = (projection || []).map(item => [
    normalizeNopol_(item.nopol),
    parseCalibrationNumber_(item.calibrationNumber),
    normalizeLookupText_(item.statusUnit),
    normalizeLookupText_(item.status3L)
  ]);
  return sha256Hex_(JSON.stringify(state));
}

/**
 * Baseline watcher TANPA mengubah data existing.
 * Hanya menyimpan fingerprint kondisi MASTER_UNIT saat trigger dipasang/di-repair.
 * Tujuan: watcher mulai mengamati perubahan SETELAH baseline, bukan merekonsiliasi
 * drift lama yang sudah ada sebelum versi watcher dipasang.
 */
function initializeStatus3LWatchBaseline_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  const now = new Date();
  const result = { initialized: false, scanned: 0, signature: '', at: now.toISOString() };
  if (!sheet || sheet.getLastRow() < 2 || !sheet.getLastColumn()) {
    PropertiesService.getScriptProperties().setProperty(KM_STATUS_3L_WATCH.LAST_RUN_PROPERTY, result.at);
    return result;
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  if ([indexes.nopol, indexes.statusUnit, indexes.calibrationNumber, indexes.status3L].some(index => index < 0)) {
    PropertiesService.getScriptProperties().setProperty(KM_STATUS_3L_WATCH.LAST_RUN_PROPERTY, result.at);
    return result;
  }
  const projection = status3LWatchBusinessProjection_(sheet, indexes);
  const signature = status3LWatchSignature_(projection);
  PropertiesService.getScriptProperties().setProperties({
    [KM_STATUS_3L_WATCH.LAST_SIGNATURE_PROPERTY]: signature,
    [KM_STATUS_3L_WATCH.LAST_RUN_PROPERTY]: result.at
  });
  result.initialized = true;
  result.scanned = projection.length;
  result.signature = signature;
  return result;
}

function syncExternalStatus3LChangesCore_(options) {
  const settings = options || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  const report = {
    success: true,
    scanned: 0,
    businessStateChanged: false,
    reconciled: 0,
    finalSnapshots: 0,
    readyToDepart: 0,
    invalidRestored: 0,
    skippedRecentManual: 0,
    unresolved: 0,
    samples: []
  };
  if (!sheet || sheet.getLastRow() < 2 || !sheet.getLastColumn()) return report;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  const requiredFields = ['nopol','statusUnit','calibrationNumber','status3L',
    'status3LUpdatedAt','status3LUpdatedBy','status3LNote'];
  if (requiredFields.some(field => indexes[field] < 0)) return report;

  const projection = status3LWatchBusinessProjection_(sheet, indexes);
  report.scanned = projection.length;
  const signature = status3LWatchSignature_(projection);
  const props = PropertiesService.getScriptProperties();
  const previousSignature = props.getProperty(KM_STATUS_3L_WATCH.LAST_SIGNATURE_PROPERTY) || '';
  report.businessStateChanged = Boolean(previousSignature) && signature !== previousSignature;

  // Safety v3: instalasi / repair trigger TIDAK BOLEH menyentuh data existing.
  // Jika watcher belum memiliki baseline, jadikan kondisi saat ini sebagai titik awal
  // dan berhenti. Hanya perubahan setelah baseline yang boleh diproses otomatis.
  if (!previousSignature && !settings.forceExistingRepair) {
    props.setProperties({
      [KM_STATUS_3L_WATCH.LAST_SIGNATURE_PROPERTY]: signature,
      [KM_STATUS_3L_WATCH.LAST_RUN_PROPERTY]: new Date().toISOString()
    });
    report.baselineInitialized = true;
    report.baselineOnly = true;
    return report;
  }

  if (!settings.force && previousSignature && signature === previousSignature) {
    props.setProperty(KM_STATUS_3L_WATCH.LAST_RUN_PROPERTY, new Date().toISOString());
    return report;
  }

  // Metadata audit (UPDATED AT/BY/NOTE) baru dibaca ketika business signature berubah.
  // Pada mayoritas tick watcher, request berhenti di atas sehingga beban Sheet jauh lebih kecil.
  status3LWatchAttachMetadata_(sheet, indexes, projection);

  const status3LLog = ss.getSheetByName(UNIT_STATUS_3L_LOG_SHEET);
  const trustedIndex = buildTrustedStatus3LStateIndex_(status3LLog);
  const now = new Date();
  const statusEntries = [];
  const unitEntries = [];
  const finalEntries = [];
  const changedRows = [];
  const postProjection = projection.map(item => Object.assign({}, item));

  projection.forEach((item, offset) => {
    const sequence = parseCalibrationNumber_(item.calibrationNumber);
    if (!sequence) return;
    const key = calibrationCompletionKey_(item.nopol, sequence);
    const bucket = key ? trustedIndex[key] || null : null;
    const latest = bucket && bucket.latest || null;
    const latestStatus = latest ? displayValue_(latest.status3L) : '';
    const currentStatus = displayValue_(item.status3L);
    const currentKey = normalizeLookupText_(currentStatus);
    const latestKey = normalizeLookupText_(latestStatus);
    const currentUnitKey = normalizeLookupText_(item.statusUnit);
    const blankAllowed = !currentKey &&
      currentUnitKey === normalizeLookupText_('On Delivery');

    let validCurrent = blankAllowed || !currentKey;
    if (currentKey) {
      try {
        resolveStatus3L_(currentStatus);
        validCurrent = true;
      } catch (error) {
        validCurrent = false;
      }
    }

    // Programmatic write can bypass data validation. Restore only from a trusted
    // existing event; never invent a replacement for an invalid value.
    if (!validCurrent) {
      if (latest && cleanText_(latestStatus)) {
        const live = sheet.getRange(item.rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
        const liveStatus = displayValue_(live[indexes.status3L]);
        if (normalizeLookupText_(liveStatus) !== currentKey) return;
        const restoreAt = latest.timestamp || now;
        const restoreBy = normalizeDoneStatus3LActor_(latestStatus, latest.updatedBy || 'SYSTEM');
        sheet.getRange(item.rowNumber, indexes.status3L + 1, 1, 4).setValues([[
          safeSheetText_(latestStatus), restoreAt, safeSheetText_(restoreBy),
          safeSheetText_(latest.note || item.status3LNote)
        ]]);
        postProjection[offset].status3L = latestStatus;
        postProjection[offset].status3LUpdatedAt = restoreAt;
        postProjection[offset].status3LUpdatedBy = restoreBy;
        if (normalizeLookupText_(latestStatus) === 'DONE' &&
            normalizeLookupText_(item.statusUnit) !== normalizeLookupText_('Ready to Depart')) {
          sheet.getRange(item.rowNumber, indexes.statusUnit + 1).setValue('Ready to Depart');
          postProjection[offset].statusUnit = 'Ready to Depart';
          unitEntries.push({
            updateId: Utilities.getUuid(), timestamp: restoreAt, email: 'SYSTEM',
            nopol: item.nopol, calibrationNumber: sequence,
            previousStatus: item.statusUnit, statusUnit: 'Ready to Depart',
            note: 'Status Unit dipulihkan bersama trusted Status 3L DONE.',
            source: 'EXTERNAL_SYNC', reason: 'STATUS_3L_DONE'
          });
          report.readyToDepart++;
        }
        report.invalidRestored++;
        changedRows.push(item.rowNumber);
        if (report.samples.length < 100) report.samples.push({
          row: item.rowNumber, nopol: item.nopol, calibrationNumber: sequence,
          previousStatus3L: currentStatus, status3L: latestStatus,
          reason: 'INVALID_EXTERNAL_STATUS_RESTORED'
        });
      } else {
        report.unresolved++;
        if (report.samples.length < 100) report.samples.push({
          row: item.rowNumber, nopol: item.nopol, calibrationNumber: sequence,
          previousStatus3L: currentStatus, status3L: currentStatus,
          reason: 'INVALID_EXTERNAL_STATUS_WITHOUT_TRUSTED_FALLBACK'
        });
      }
      return;
    }

    // Ledger already matches current business state.
    if (currentKey === latestKey && (currentKey || blankAllowed)) return;
    if (!currentKey && !blankAllowed) {
      // Kalibrasi aktif tidak boleh tiba-tiba kehilangan Status 3L kecuali unit
      // memang On Delivery. Jangan menganggap drift ini sudah selesai.
      if (latest && cleanText_(latestStatus)) {
        report.unresolved++;
        if (report.samples.length < 100) report.samples.push({
          row: item.rowNumber, nopol: item.nopol, calibrationNumber: sequence,
          previousStatus3L: latestStatus, status3L: '', reason: 'UNEXPECTED_EXTERNAL_BLANK'
        });
      }
      return;
    }

    const existingAt = item.status3LUpdatedAt instanceof Date
      ? item.status3LUpdatedAt : parseDateTimeInput_(item.status3LUpdatedAt);
    const latestAt = latest && latest.timestamp instanceof Date ? latest.timestamp : null;
    const existingBy = status3LActorDisplay_(item.status3LUpdatedBy);
    const existingAtValid = existingAt instanceof Date && !isNaN(existingAt.getTime());

    // Manual onEdit writes UPDATED AT/BY before ledger continuation. Give it time
    // to finish so watcher does not misclassify a real editor as automation.
    if (!settings.force && existingAtValid &&
        now.getTime() - existingAt.getTime() < KM_STATUS_3L_WATCH.MANUAL_EDIT_GRACE_MS) {
      report.skippedRecentManual++;
      return;
    }
    if (!settings.force && latestAt &&
        now.getTime() - latestAt.getTime() < KM_STATUS_3L_WATCH.MANUAL_EDIT_GRACE_MS) {
      report.skippedRecentManual++;
      return;
    }

    const primaryHasNewAttribution = existingAtValid &&
      (!latestAt || existingAt.getTime() > latestAt.getTime()) &&
      isResolvedStatus3LActor_(existingBy) &&
      normalizeLookupText_(existingBy) !== 'SYSTEM';

    const detectedAt = primaryHasNewAttribution ? existingAt : now;
    const detectedBy = primaryHasNewAttribution
      ? normalizeDoneStatus3LActor_(currentStatus, existingBy)
      : normalizeDoneStatus3LActor_(currentStatus, KM_STATUS_3L_AUTOMATION_ACTOR);
    let nextUnit = item.statusUnit;

    if (currentKey === 'DONE') {
      nextUnit = 'Ready to Depart';
    }

    const live = sheet.getRange(item.rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (normalizeLookupText_(live[indexes.status3L]) !== currentKey ||
        parseCalibrationNumber_(live[indexes.calibrationNumber]) !== sequence) {
      return;
    }

    sheet.getRange(item.rowNumber, indexes.status3LUpdatedAt + 1, 1, 2)
      .setValues([[detectedAt, safeSheetText_(detectedBy)]]);
    postProjection[offset].status3LUpdatedAt = detectedAt;
    postProjection[offset].status3LUpdatedBy = detectedBy;

    if (normalizeLookupText_(nextUnit) !== currentUnitKey) {
      sheet.getRange(item.rowNumber, indexes.statusUnit + 1).setValue(safeSheetText_(nextUnit));
      postProjection[offset].statusUnit = nextUnit;
      unitEntries.push({
        updateId: Utilities.getUuid(), timestamp: detectedAt, email: 'SYSTEM',
        nopol: item.nopol, calibrationNumber: sequence,
        previousStatus: item.statusUnit, statusUnit: nextUnit,
        note: 'Ready to Depart diterapkan oleh watcher karena Status 3L eksternal terdeteksi DONE.',
        source: 'EXTERNAL_SYNC', reason: 'STATUS_3L_DONE'
      });
      report.readyToDepart++;
    }

    const reason = currentKey
      ? 'STATUS_3L_EXTERNAL_CHANGE_DETECTED'
      : 'STATUS_3L_EXTERNAL_BLANK_DETECTED';
    statusEntries.push({
      updateId: Utilities.getUuid(),
      timestamp: detectedAt,
      email: detectedBy,
      nopol: item.nopol,
      calibrationNumber: sequence,
      previousStatus: latestStatus,
      status3L: currentStatus,
      note: item.status3LNote ||
        'Perubahan Status 3L terdeteksi dari MASTER_UNIT karena external write/automation tidak memicu onEdit.',
      source: 'EXTERNAL_SYNC',
      reason: reason,
      previousStatusUnit: item.statusUnit,
      statusUnit: nextUnit
    });

    if (currentKey === 'DONE') {
      const finalRow = live.slice();
      finalRow[indexes.statusUnit] = nextUnit;
      finalRow[indexes.status3LUpdatedAt] = detectedAt;
      finalRow[indexes.status3LUpdatedBy] = detectedBy;
      finalEntries.push({
        updateId: Utilities.getUuid(), timestamp: detectedAt, email: detectedBy,
        nopol: item.nopol, calibrationNumber: sequence,
        status3L: 'DONE', source: 'EXTERNAL_SYNC', reason: 'STATUS_3L_DONE',
        dataAfter: unitKmSnapshot_(finalRow, indexes)
      });
    }

    report.reconciled++;
    changedRows.push(item.rowNumber);
    if (report.samples.length < 100) report.samples.push({
      row: item.rowNumber,
      nopol: item.nopol,
      calibrationNumber: sequence,
      previousStatus3L: latestStatus,
      status3L: currentStatus,
      updatedAt: detectedAt.toISOString(),
      updatedBy: detectedBy,
      source: 'EXTERNAL_SYNC'
    });
  });

  if (statusEntries.length) appendStatus3LHistoryRows_(statusEntries);
  if (unitEntries.length) appendStatusUnitHistoryRows_(unitEntries);
  if (finalEntries.length) {
    const finalResult = appendCompletedCalibrationSnapshots_(finalEntries);
    report.finalSnapshots = finalResult && Number(finalResult.count) ? Number(finalResult.count) : 0;
  }

  if (changedRows.length) {
    SpreadsheetApp.flush();
    const uniqueRows = [...new Set(changedRows)].sort((a, b) => a - b);
    const analyticsStates = uniqueRows.map(rowNumber => {
      const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
      return analyticsStateFromMasterRow_(row, indexes);
    }).filter(Boolean);
    if (analyticsStates.length) {
      runNonCritical_('sync analytics dari watcher Status 3L', () =>
        updateAnalyticsCalibrationState_(analyticsStates));
    }
    invalidateAppCaches_();
  }

  const propertiesToWrite = {
    [KM_STATUS_3L_WATCH.LAST_RUN_PROPERTY]: now.toISOString()
  };
  // Jika event manual masih dalam grace period atau ada drift yang belum bisa
  // direkonsiliasi, pertahankan signature lama agar tick berikutnya memeriksa lagi.
  // Ini mencegah satu race condition membuat perubahan selamanya tidak terdeteksi.
  if (!report.skippedRecentManual && !report.unresolved) {
    propertiesToWrite[KM_STATUS_3L_WATCH.LAST_SIGNATURE_PROPERTY] = status3LWatchSignature_(postProjection);
  }
  props.setProperties(propertiesToWrite);
  return report;
}

function syncExternalStatus3LChanges(e) {
  // Manual Run dari Apps Script editor tidak boleh mengubah data existing.
  // Time-driven trigger mengirim event/triggerUid; tanpa event kita hanya memastikan baseline.
  if (!e || !e.triggerUid) {
    return {
      success: true,
      manualRun: true,
      dataChanged: false,
      baseline: ensureStatus3LWatchBaseline_(),
      message: 'Manual Run = baseline only. Tidak ada data existing yang diubah.'
    };
  }
  let lock = null;
  const props = PropertiesService.getScriptProperties();
  try {
    lock = acquireScriptLock_(25000, 'watcher Status 3L eksternal');
    const result = syncExternalStatus3LChangesCore_({ force: false });
    props.setProperties({
      [KM_STATUS_3L_WATCH.LAST_SUCCESS_PROPERTY]: new Date().toISOString(),
      [KM_STATUS_3L_WATCH.LAST_ERROR_PROPERTY]: '',
      [KM_STATUS_3L_WATCH.ERROR_COUNT_PROPERTY]: '0'
    });
    return result;
  } catch (error) {
    const count = Math.max(0, Number(props.getProperty(KM_STATUS_3L_WATCH.ERROR_COUNT_PROPERTY) || 0)) + 1;
    props.setProperties({
      [KM_STATUS_3L_WATCH.LAST_ERROR_PROPERTY]: String(error && error.message ? error.message : error).slice(0, 500),
      [KM_STATUS_3L_WATCH.ERROR_COUNT_PROPERTY]: String(count)
    });
    throw error;
  } finally {
    if (lock) lock.releaseLock();
  }
}

function ensureStatus3LWatchBaseline_() {
  const props = PropertiesService.getScriptProperties();
  const existing = props.getProperty(KM_STATUS_3L_WATCH.LAST_SIGNATURE_PROPERTY) || '';
  if (existing) {
    return { initialized: false, preserved: true, signature: existing,
      at: props.getProperty(KM_STATUS_3L_WATCH.LAST_RUN_PROPERTY) || '' };
  }
  return initializeStatus3LWatchBaseline_();
}

function auditExternalStatus3LDriftFromSheet() {
  requireSheetMaintenanceAdmin_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  const report = { scanned: 0, drift: 0, invalid: 0, blankUnexpected: 0, samples: [] };
  if (!sheet || sheet.getLastRow() < 2 || !sheet.getLastColumn()) return report;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  if (indexes.nopol < 0 || indexes.statusUnit < 0 || indexes.calibrationNumber < 0 || indexes.status3L < 0) return report;
  const projection = status3LWatchBusinessProjection_(sheet, indexes);
  const trusted = buildTrustedStatus3LStateIndex_(ss.getSheetByName(UNIT_STATUS_3L_LOG_SHEET));
  report.scanned = projection.length;
  projection.forEach(item => {
    const sequence = parseCalibrationNumber_(item.calibrationNumber);
    if (!sequence) return;
    const key = calibrationCompletionKey_(item.nopol, sequence);
    const latest = key && trusted[key] ? trusted[key].latest : null;
    const current = displayValue_(item.status3L);
    const currentKey = normalizeLookupText_(current);
    const latestStatus = latest ? displayValue_(latest.status3L) : '';
    const latestKey = normalizeLookupText_(latestStatus);
    let valid = !currentKey;
    if (currentKey) {
      try { resolveStatus3L_(current); valid = true; } catch (error) { valid = false; }
    }
    const unexpectedBlank = !currentKey && normalizeLookupText_(item.statusUnit) !== normalizeLookupText_('On Delivery') && cleanText_(latestStatus);
    const drift = !valid || unexpectedBlank || (currentKey && currentKey !== latestKey);
    if (!drift) return;
    report.drift++;
    if (!valid) report.invalid++;
    if (unexpectedBlank) report.blankUnexpected++;
    if (report.samples.length < 200) report.samples.push({
      row: item.rowNumber, nopol: item.nopol, calibrationNumber: sequence,
      historyStatus3L: latestStatus, masterStatus3L: current,
      statusUnit: item.statusUnit, invalid: !valid, unexpectedBlank: unexpectedBlank
    });
  });
  return report;
}

function repairExternalStatus3LChangesFromSheet(confirmation) {
  const actor = requireSheetMaintenanceAdmin_();
  // Safety v3: menjalankan fungsi dari tombol Run Apps Script tanpa parameter
  // tidak akan mengubah MASTER_UNIT/history. Hanya menampilkan audit drift.
  if (cleanText_(confirmation) !== 'REPAIR_EXISTING_STATUS_3L') {
    const audit = auditExternalStatus3LDriftFromSheet();
    return {
      success: true,
      dryRun: true,
      dataChanged: false,
      confirmationRequired: 'REPAIR_EXISTING_STATUS_3L',
      audit: audit,
      message: 'DRY RUN: tidak ada data existing yang diubah. Repair existing hanya tersedia dengan confirmation eksplisit.'
    };
  }
  KM_REQUEST_ACTOR_ = actor;
  let lock = null;
  try {
    lock = acquireScriptLock_(60000, 'repair Status 3L eksternal');
    const result = syncExternalStatus3LChangesCore_({ force: true, forceExistingRepair: true });
    logSystemActivity_('', '', 'STATUS_3L_EXTERNAL_SYNC_REPAIR', 'MASTER_UNIT', '', '',
      JSON.stringify({
        reconciled: result.reconciled,
        invalidRestored: result.invalidRestored,
        finalSnapshots: result.finalSnapshots
      }), actor, true);
    result.dataChanged = Boolean(result.reconciled || result.invalidRestored || result.readyToDepart || result.finalSnapshots);
    return result;
  } finally {
    if (lock) lock.releaseLock();
  }
}

function ensureStatus3LWatchTrigger_() {
  // Baseline hanya dibuat bila belum ada. Sinkronisasi trigger berikutnya TIDAK
  // mereset signature karena reset baseline dapat menelan perubahan yang belum sempat
  // diproses watcher. Fungsi ini tetap tidak mengubah data bisnis existing.
  const baseline = ensureStatus3LWatchBaseline_();
  const candidates = ScriptApp.getProjectTriggers().filter(trigger =>
    trigger.getHandlerFunction() === KM_STATUS_3L_WATCH.HANDLER);
  let created = null;
  if (!candidates.length) {
    created = ScriptApp.newTrigger(KM_STATUS_3L_WATCH.HANDLER)
      .timeBased()
      .everyMinutes(KM_STATUS_3L_WATCH.INTERVAL_MINUTES)
      .create();
  }
  const canonical = created || candidates[0];
  const duplicates = created ? candidates : candidates.slice(1);
  duplicates.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  PropertiesService.getScriptProperties().setProperties({
    [KM_STATUS_3L_WATCH.ENABLED_PROPERTY]: 'TRUE',
    [KM_STATUS_3L_WATCH.ID_PROPERTY]: String(canonical.getUniqueId())
  });
  return {
    installed: Boolean(created),
    ready: true,
    intervalMinutes: KM_STATUS_3L_WATCH.INTERVAL_MINUTES,
    existingBefore: candidates.length,
    removedDuplicates: duplicates.length,
    handler: KM_STATUS_3L_WATCH.HANDLER,
    baseline: baseline,
    existingDataChanged: false
  };
}

function isStatus3LWatchTriggerReady_() {
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty(KM_STATUS_3L_WATCH.ENABLED_PROPERTY) !== 'TRUE') return false;
    const canonicalId = props.getProperty(KM_STATUS_3L_WATCH.ID_PROPERTY);
    const triggers = ScriptApp.getProjectTriggers().filter(trigger =>
      trigger.getHandlerFunction() === KM_STATUS_3L_WATCH.HANDLER &&
      trigger.getEventType() === ScriptApp.EventType.CLOCK
    );
    return triggers.length === 1 &&
      (!canonicalId || String(triggers[0].getUniqueId()) === canonicalId);
  } catch (error) {
    return false;
  }
}

function ensureKalibrasiKmEditTrigger_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const candidates = ScriptApp.getProjectTriggers().filter(trigger =>
    trigger.getHandlerFunction() === KM_SHEET_EDIT_TRIGGER.HANDLER);
  const triggers = candidates.filter(trigger => isKalibrasiKmEditTriggerForSpreadsheet_(trigger, spreadsheet.getId()));
  let created = null;
  if (!triggers.length) {
    created = ScriptApp.newTrigger(KM_SHEET_EDIT_TRIGGER.HANDLER)
      .forSpreadsheet(spreadsheet)
      .onEdit()
      .create();
  }
  // Preserve one existing trigger instead of deleting/recreating every repair.
  const duplicates = triggers.slice(1);
  duplicates.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  const canonical = created || triggers[0];
  PropertiesService.getScriptProperties().setProperties({
    [KM_SHEET_EDIT_TRIGGER.ENABLED_PROPERTY]: 'TRUE',
    [KM_SHEET_EDIT_TRIGGER.ID_PROPERTY]: String(canonical.getUniqueId())
  });
  return {
    installed: Boolean(created),
    ready: true,
    existingBefore: triggers.length,
    removedDuplicates: duplicates.length,
    ignoredOtherTriggers: candidates.length - triggers.length,
    handler: created ? created.getHandlerFunction() : KM_SHEET_EDIT_TRIGGER.HANDLER
  };
}

function isKalibrasiKmEditTriggerForSpreadsheet_(trigger, spreadsheetId) {
  return trigger.getHandlerFunction() === KM_SHEET_EDIT_TRIGGER.HANDLER &&
    trigger.getEventType() === ScriptApp.EventType.ON_EDIT &&
    trigger.getTriggerSourceId() === spreadsheetId;
}

function isKalibrasiKmEditTriggerReady_() {
  try {
    const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    const properties = PropertiesService.getScriptProperties();
    if (properties.getProperty(KM_SHEET_EDIT_TRIGGER.ENABLED_PROPERTY) !== 'TRUE') return false;
    const canonicalId = properties.getProperty(KM_SHEET_EDIT_TRIGGER.ID_PROPERTY);
    const triggers = ScriptApp.getProjectTriggers().filter(trigger =>
      isKalibrasiKmEditTriggerForSpreadsheet_(trigger, spreadsheetId)
    );
    return triggers.length === 1 &&
      (!canonicalId || String(triggers[0].getUniqueId()) === canonicalId);
  } catch (error) {
    return false;
  }
}