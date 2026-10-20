/** Administration Center, master-data governance, import/export, and maintenance. */
const KM_ADMIN_BACKUP_FOLDER_PROPERTY = 'KM_ADMIN_BACKUP_FOLDER_ID_V1';
const KM_ADMIN_BACKUP_FOLDER_NAME = 'Kalibrasi KM Backups';
const KM_ADMIN_BACKUP_PREFIX = 'KALIBRASI_KM_BACKUP_';
const KM_ADMIN_IMPORT_MAX_BYTES = 2 * 1024 * 1024;
const KM_ADMIN_IMPORT_MAX_ROWS = 1500;
// Literal sheet names keep this top-level constant independent from file load order.
const KM_ADMIN_RESTORE_SHEETS = Object.freeze([
  'MASTER_CONFIG',
  'MASTER_UNIT',
  'KM_USERS',
  'KM_ROLES',
  'KM_ACTIVITY_LOG',
  'HISTORY_KALIBRASI_KM',
  'ANALYTICS_KALIBRASI_KM',
  'STATUS_UNIT_HISTORY',
  'STATUS_3L_HISTORY',
  'MASTER_UNIT_MOVEMENT_HISTORY',
  'MASTER_UNIT_CHANGE_REQUESTS'
]);

const MASTER_UNIT_MANAGEMENT = Object.freeze({
  MOVEMENT_SHEET: 'MASTER_UNIT_MOVEMENT_HISTORY',
  REQUEST_SHEET: 'MASTER_UNIT_CHANGE_REQUESTS',
  MASTER_STATUSES: ['ACTIVE', 'INACTIVE'],
  APPROVAL_PARAMETER: 'MASTER_UNIT_APPROVAL_ENABLED',
  BULK_MAX_ROWS: 1000,
  MOVEMENT_HEADERS: [
    'MOVEMENT ID','TIMESTAMP','UNIT ID','NOMOR POLISI','MOVEMENT TYPE',
    'PREVIOUS OPERATING POINT','NEW OPERATING POINT','PREVIOUS TOH/WOH','NEW TOH/WOH',
    'PREVIOUS TYPE','NEW TYPE','PREVIOUS NOMOR POLISI','NEW NOMOR POLISI',
    'EFFECTIVE DATE','REASON','NOTE','UPDATED BY','SOURCE','REQUEST ID'
  ],
  REQUEST_HEADERS: [
    'REQUEST ID','TIMESTAMP','REQUESTED BY','REQUESTED ROLE','ACTION','UNIT ID','NOMOR POLISI',
    'CURRENT DATA JSON','PROPOSED DATA JSON','EFFECTIVE DATE','REASON','NOTE','STATUS',
    'REVIEWED AT','REVIEWED BY','REVIEW NOTE','APPLIED AT','APPLIED BY','MOVEMENT ID'
  ]
});

const MASTER_UNIT_ADMIN_HEADERS = Object.freeze({
  unitId: 'UNIT ID',
  masterStatus: 'MASTER STATUS',
  masterCreatedAt: 'MASTER CREATED AT',
  masterCreatedBy: 'MASTER CREATED BY',
  masterUpdatedAt: 'MASTER UPDATED AT',
  masterUpdatedBy: 'MASTER UPDATED BY',
  lastMovementId: 'LAST MOVEMENT ID'
});

const MASTER_UNIT_DIRECTORY_HEADERS = Object.freeze({
  unitType: 'TIPE'
});

function requireAdminCenterPermission_(permission, sessionToken) {
  const actor = requirePermission_(permission, '', { token: sessionToken });
  KM_REQUEST_ACTOR_ = actor;
  return actor;
}

function getAdminAreaCoverageData(sessionToken) {
  requireAdminCenterPermission_('ADMIN_SETUP', sessionToken);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.CONFIG);
  if (!sheet) throw new Error('MASTER_CONFIG belum tersedia.');
  const coverage = masterConfigCoverageDirectory_();
  return {
    revision: coverageDirectoryRevision_(coverage),
    rows: coverage.map(item => ({
      operatingPoint: item.operatingPoint || '',
      opc: item.opc || '',
      toh: item.toh || ''
    })),
    stats: {
      rows: coverage.length,
      operatingPoints: uniqueClean_(coverage.map(item => item.operatingPoint)).length,
      opc: uniqueClean_(coverage.map(item => item.opc)).length,
      toh: uniqueClean_(coverage.map(item => item.toh)).length
    }
  };
}

function normalizeAdminCoverageRows_(rows) {
  if (!Array.isArray(rows)) throw new Error('Format Area & Coverage tidak valid.');
  if (rows.length > KM_ADMIN_IMPORT_MAX_ROWS) throw new Error('Maksimal ' + KM_ADMIN_IMPORT_MAX_ROWS + ' baris per penyimpanan.');
  const seen = new Set();
  return rows.map((item, index) => ({
    operatingPoint: cleanText_(item && item.operatingPoint),
    opc: cleanText_(item && item.opc),
    toh: cleanText_(item && item.toh),
    sourceRow: index + 1
  })).filter(item => item.operatingPoint || item.opc || item.toh)
    .map(item => {
      if (!item.operatingPoint) throw new Error('Operating Point wajib diisi pada baris ' + item.sourceRow + '.');
      const key = [item.operatingPoint, item.opc, item.toh].map(normalizeLookupText_).join('|');
      if (seen.has(key)) throw new Error('Duplikat Area & Coverage pada baris ' + item.sourceRow + '.');
      seen.add(key);
      return item;
    });
}

function saveAdminAreaCoverage(rows, sessionToken, expectedRevision) {
  const actor = requireAdminCenterPermission_('ADMIN_SETUP', sessionToken);
  const normalized = normalizeAdminCoverageRows_(rows);
  const lock = acquireScriptLock_(30000, 'simpan Area & Coverage');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet_(ss, KM_APP.SHEETS.CONFIG);
    if (expectedRevision && expectedRevision !== coverageDirectoryRevision_(masterConfigCoverageDirectory_())) {
      throw new Error('DATA_CONFLICT: Master Config sudah diubah pengguna lain. Muat ulang sebelum menyimpan.');
    }
    if (sheet.getMaxRows() < 2) sheet.insertRowsAfter(1, 1);
    if (sheet.getMaxColumns() < 6) sheet.insertColumnsAfter(sheet.getMaxColumns(), 6 - sheet.getMaxColumns());
    sheet.getRange('A1:C1').setValues([['OPERATING POINT', 'NAMA OPC/WHC', 'NAMA TOH/WOH']]);
    const previousRows = Math.max(sheet.getLastRow() - 1, 0);
    if (normalized.length) {
      if (sheet.getMaxRows() < normalized.length + 1) {
        sheet.insertRowsAfter(sheet.getMaxRows(), normalized.length + 1 - sheet.getMaxRows());
      }
      sheet.getRange(2, 1, normalized.length, 3).setValues(normalized.map(item => [
        safeSheetText_(item.operatingPoint), safeSheetText_(item.opc), safeSheetText_(item.toh)
      ]));
    }
    if (previousRows > normalized.length) sheet.getRange(normalized.length + 2, 1, previousRows - normalized.length, 3).clearContent();
    SpreadsheetApp.flush();
    invalidateAppCaches_();
    logSystemActivity_('', '', 'ADMIN_AREA_COVERAGE_UPDATED', 'MASTER_CONFIG A:C', '',
      String(normalized.length) + ' baris', 'Area & Coverage diperbarui dari Administration.', actor, true);
    return getAdminAreaCoverageData(sessionToken);
  } finally {
    lock.releaseLock();
  }
}

function getAdminWorkflowConfig(sessionToken) {
  requireAdminCenterPermission_('ADMIN_SETUP', sessionToken);
  const params = getConfigParameters_();
  return {
    appName: getConfigParameter_('APP_NAME', 'Monitoring Kalibrasi KM'),
    pageSize: Number(getConfigParameter_('PAGE_SIZE', 15)) || 15,
    schemaVersion: Number(getConfigParameter_('SCHEMA_VERSION', KM_DATABASE_SCHEMA_VERSION)) || KM_DATABASE_SCHEMA_VERSION,
    codeSchemaVersion: KM_DATABASE_SCHEMA_VERSION,
    status3LOptions: getStatus3LOptions_(),
    requiredStatus3L: DEFAULT_STATUS_3L_OPTIONS.slice(),
    manualUnitStatuses: UNIT_MANUAL_OPERATIONAL_STATUSES.slice(),
    operationalStatuses: UNIT_OPERATIONAL_STATUSES.slice(),
    masterUnitApprovalEnabled: masterUnitApprovalEnabled_(),
    submissionSchedule: getCalibrationSubmissionSchedule_(),
    editTriggerReady: isKalibrasiKmEditTriggerReady_(),
    status3LWatchReady: isStatus3LWatchTriggerReady_(),
    parameters: params
  };
}

function saveAdminWorkflowConfig(payload, sessionToken) {
  const actor = requireAdminCenterPermission_('ADMIN_SETUP', sessionToken);
  const data = payload || {};
  const appName = cleanText_(data.appName) || 'Monitoring Kalibrasi KM';
  if (appName.length > 100) throw new Error('Nama aplikasi maksimal 100 karakter.');
  const pageSize = Number(data.pageSize);
  if (![15, 30, 50].includes(pageSize)) throw new Error('PAGE_SIZE hanya boleh 15, 30, atau 50.');
  const approvalEnabled = masterUnitApprovalEnabled_();
  const existingParameters = getConfigParameters_();
  const submissionSchedule = normalizeCalibrationSubmissionSchedule_(
    data.submissionSchedule === undefined
      ? calibrationSubmissionScheduleFromParameters_(existingParameters)
      : data.submissionSchedule
  );
  const status3L = uniqueClean_(DEFAULT_STATUS_3L_OPTIONS.concat(
    Array.isArray(data.status3LOptions) ? data.status3LOptions : []
  ));
  if (status3L.length > 30) throw new Error('Maksimal 30 opsi Status 3L.');
  status3L.forEach(status => {
    if (status.length > 60) throw new Error('Status 3L terlalu panjang: ' + status);
  });

  const lock = acquireScriptLock_(30000, 'simpan Workflow Configuration');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet_(ss, KM_APP.SHEETS.CONFIG);
    if (sheet.getMaxColumns() < 6) sheet.insertColumnsAfter(sheet.getMaxColumns(), 6 - sheet.getMaxColumns());
    const parameterOverrides = Object.assign({
      SCHEMA_VERSION: String(KM_DATABASE_SCHEMA_VERSION),
      PAGE_SIZE: pageSize,
      APP_NAME: appName,
      [MASTER_UNIT_MANAGEMENT.APPROVAL_PARAMETER]: approvalEnabled ? 'TRUE' : 'FALSE'
    }, calibrationSubmissionScheduleParameterOverrides_(submissionSchedule));
    const parameterRows = buildManagedConfigParameterRows_(existingParameters, parameterOverrides);
    const requiredRows = Math.max(status3L.length, parameterRows.length) + 1;
    if (sheet.getMaxRows() < requiredRows) {
      sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
    }
    sheet.getRange('D1:F1').setValues([['STATUS PROSES 3L', 'PARAMETER', 'VALUE']]);
    sheet.getRange(2, 4, Math.max(sheet.getMaxRows() - 1, 1), 3).clearContent();
    if (status3L.length) sheet.getRange(2, 4, status3L.length, 1).setValues(status3L.map(value => [safeSheetText_(value)]));
    if (parameterRows.length) sheet.getRange(2, 5, parameterRows.length, 2).setValues(parameterRows);
    const master = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
    if (master && master.getLastColumn()) {
      const indexes = buildAliasIndex_(master.getRange(1, 1, 1, master.getLastColumn()).getDisplayValues()[0], MASTER_UNIT_FIELD_ALIASES);
      applyStatus3LValidation_(master, indexes);
    }
    SpreadsheetApp.flush();
    invalidateAppCaches_();
    logSystemActivity_('', '', 'ADMIN_WORKFLOW_CONFIG_UPDATED', 'MASTER_CONFIG D:F', '',
      JSON.stringify({ appName: appName, pageSize: pageSize, status3L: status3L,
        submissionSchedule: submissionSchedule }),
      'Workflow Configuration diperbarui.', actor, true);
    return getAdminWorkflowConfig(sessionToken);
  } finally {
    lock.releaseLock();
  }
}

function getAdminMasterUnitOverview(sessionToken) {
  requireAdminCenterPermission_('ADMIN_SETUP', sessionToken);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  if (!sheet) throw new Error('MASTER_UNIT belum tersedia.');
  const headers = sheet.getLastColumn()
    ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(cleanText_)
    : [];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  const counts = { statusUnit: {}, status3L: {}, calibrated: 0, uncalibrated: 0 };
  if (sheet.getLastRow() > 1) {
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
    values.forEach(row => {
      if (indexes.nopol < 0 || !cleanText_(row[indexes.nopol])) return;
      const su = cleanText_(row[indexes.statusUnit]) || '(KOSONG)';
      const s3 = cleanText_(row[indexes.status3L]) || '(KOSONG)';
      counts.statusUnit[su] = (counts.statusUnit[su] || 0) + 1;
      counts.status3L[s3] = (counts.status3L[s3] || 0) + 1;
      if (parseCalibrationNumber_(row[indexes.calibrationNumber])) counts.calibrated++;
      else counts.uncalibrated++;
    });
  }
  return {
    spreadsheetName: ss.getName(),
    rows: Math.max(0, sheet.getLastRow() - 1),
    columns: sheet.getLastColumn(),
    maxRows: sheet.getMaxRows(),
    maxColumns: sheet.getMaxColumns(),
    headers: headers,
    counts: counts,
    audit: auditMasterUnitKmDatabase_()
  };
}

function repairAdminMasterUnit(sessionToken, deferOverview) {
  const actor = requireAdminCenterPermission_('ADMIN_SETUP', sessionToken);
  const result = setupMasterUnitKmDatabase_();
  // Repair Struktur juga menjalankan invariant deterministik DONE -> Ready to Depart.
  // Ini tidak mengubah Status 3L dan tidak menebak mapping coverage.
  const doneReadyInvariant = repairStatus3LDoneReadyInvariant_(actor);
  // v17: sinkronkan kembali ID foto yang sebenarnya sudah tersimpan pada ledger
  // HISTORY_KALIBRASI_KM tetapi kolom FOTO ODOMETER ID di MASTER_UNIT kosong.
  // Repair bersifat idempotent dan hanya mengisi field kosong jika kandidat unik.
  const photoIdRecovery = repairOdometerPhotoIdsFromHistory_(actor);
  logSystemActivity_('', '', 'ADMIN_MASTER_UNIT_REPAIR', 'MASTER_UNIT', '', '',
    'Repair struktur MASTER_UNIT + invariant DONE -> Ready to Depart + sinkronisasi ID foto odometer dari history.', actor);
  return {
    success: true,
    result: result,
    doneReadyInvariant: doneReadyInvariant,
    photoIdRecovery: photoIdRecovery,
    overviewDeferred: deferOverview === true,
    overview: deferOverview === true ? null : getAdminMasterUnitOverview(sessionToken),
    management: deferOverview === true ? null : getAdminMasterUnitManagementData(sessionToken)
  };
}

function repairAdminMasterUnitDataQuality(sessionToken) {
  const actor = requireAdminCenterPermission_('ADMIN_SETUP', sessionToken);
  const lock = acquireScriptLock_(30000, 'repair Data Quality Status 3L');
  let doneReadyInvariant;
  try { doneReadyInvariant = repairStatus3LDoneReadyInvariant_(actor); }
  finally { lock.releaseLock(); }
  const management = getAdminMasterUnitManagementData(sessionToken);
  logSystemActivity_('', '', 'ADMIN_MASTER_DATA_QUALITY_REPAIR', 'MASTER_UNIT', '', '',
    'Repair aman Data Quality: invariant DONE -> Ready to Depart. Coverage tidak diubah otomatis.', actor, true);
  return {
    success: true,
    doneReadyInvariant: doneReadyInvariant,
    management: management,
    quality: management.quality
  };
}

function requireAnyMasterUnitManagementPermission_(sessionToken) {
  const actor = getCurrentUserContext_({ token: sessionToken });
  const allowed = ['MANAGE_MASTER_UNIT', 'REQUEST_MASTER_UNIT_CHANGE', 'APPROVE_MASTER_UNIT_CHANGE'];
  if (!allowed.some(permission => actor.permissions.includes(permission))) {
    throw new Error('ACCESS_DENIED: akun tidak memiliki akses Master Unit Management.');
  }
  KM_REQUEST_ACTOR_ = actor;
  return actor;
}

function masterUnitApprovalEnabled_() {
  const value = normalizeLookupText_(getConfigParameter_(MASTER_UNIT_MANAGEMENT.APPROVAL_PARAMETER, 'FALSE'));
  return ['TRUE', '1', 'YES', 'ON', 'ENABLED'].includes(value);
}

function newMasterUnitId_() {
  return 'UNIT-' + Utilities.getUuid().replace(/-/g, '').slice(0, 12).toUpperCase();
}

function newMasterUnitMovementId_(timestamp) {
  const date = timestamp instanceof Date ? timestamp : new Date();
  return 'MOV-' + Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' +
    Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase();
}

function newMasterUnitRequestId_(timestamp) {
  const date = timestamp instanceof Date ? timestamp : new Date();
  return 'MUR-' + Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' +
    Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase();
}

function formatMasterUnitManagementSheet_(sheet, headers, headerColor) {
  if (!sheet) return;
  const map = ensureHistoryHeaders_(sheet, headers);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .setBackground(headerColor || '#16324F').setFontColor('#FFFFFF').setFontWeight('bold');
  ['TIMESTAMP','REVIEWED AT','APPLIED AT'].forEach(header => {
    if (map[header]) trySetNumberFormat_(sheet.getRange(2, map[header], Math.max(sheet.getMaxRows() - 1, 1), 1), 'dd mmmm yyyy, HH.mm.ss');
  });
  if (map['EFFECTIVE DATE']) trySetNumberFormat_(sheet.getRange(2, map['EFFECTIVE DATE'], Math.max(sheet.getMaxRows() - 1, 1), 1), 'dd mmmm yyyy');
  if (map['NOTE']) sheet.setColumnWidth(map['NOTE'], 280);
  if (map['REASON']) sheet.setColumnWidth(map['REASON'], 220);
}

function masterUnitManagementSchemaReady_(sheet) {
  if (!sheet || !sheet.getLastColumn()) return false;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  return Object.keys(MASTER_UNIT_ADMIN_HEADERS).concat(Object.keys(MASTER_UNIT_DIRECTORY_HEADERS))
    .every(field => indexes[field] >= 0);
}

function ensureMasterUnitManagementSchema_(ss, masterSheet, actor) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = masterSheet || getOrCreateSheet_(spreadsheet, KM_APP.SHEETS.MASTER_UNIT);
  if (sheet.getMaxRows() < 2) sheet.insertRowsAfter(1, 1);
  let headers = sheet.getLastColumn()
    ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(cleanText_) : [];
  let indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  const added = [];
  Object.keys(MASTER_UNIT_DIRECTORY_HEADERS).forEach(field => {
    if (indexes[field] >= 0) return;
    const header = MASTER_UNIT_DIRECTORY_HEADERS[field];
    if (sheet.getMaxColumns() < headers.length + 1) sheet.insertColumnAfter(sheet.getMaxColumns());
    sheet.getRange(1, headers.length + 1).setValue(header);
    headers.push(header); added.push(header);
    indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  });
  Object.keys(MASTER_UNIT_ADMIN_HEADERS).forEach(field => {
    if (indexes[field] >= 0) return;
    const header = MASTER_UNIT_ADMIN_HEADERS[field];
    if (sheet.getMaxColumns() < headers.length + 1) sheet.insertColumnAfter(sheet.getMaxColumns());
    sheet.getRange(1, headers.length + 1).setValue(header);
    headers.push(header);
    added.push(header);
    indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  });

  // Normalisasi nama directory + metadata ke header canonical agar standardizer
  // tidak membuat kolom semantik duplikat dari alias legacy (mis. JENIS UNIT vs TIPE).
  Object.keys(MASTER_UNIT_DIRECTORY_HEADERS).forEach(field => {
    if (indexes[field] >= 0) sheet.getRange(1, indexes[field] + 1).setValue(MASTER_UNIT_DIRECTORY_HEADERS[field]);
  });
  Object.keys(MASTER_UNIT_ADMIN_HEADERS).forEach(field => {
    if (indexes[field] >= 0) sheet.getRange(1, indexes[field] + 1).setValue(MASTER_UNIT_ADMIN_HEADERS[field]);
  });

  const now = new Date();
  const migrationActor = safeSheetText_(actor && actor.email || 'SYSTEM_MIGRATION');
  let backfilledIds = 0;
  let backfilledStatus = 0;
  if (sheet.getLastRow() > 1) {
    const rowCount = sheet.getLastRow() - 1;
    const values = sheet.getRange(2, 1, rowCount, sheet.getLastColumn()).getValues();
    const unitIds = [];
    const statuses = [];
    const createdAt = [];
    const createdBy = [];
    const updatedAt = [];
    const updatedBy = [];
    const lastMovement = [];
    const seenIds = new Set();
    values.forEach(row => {
      const nopol = indexes.nopol >= 0 ? cleanText_(row[indexes.nopol]) : '';
      let unitId = cleanText_(row[indexes.unitId]);
      if (nopol && (!unitId || seenIds.has(normalizeLookupText_(unitId)))) {
        unitId = newMasterUnitId_();
        backfilledIds++;
      }
      if (unitId) seenIds.add(normalizeLookupText_(unitId));
      let status = normalizeLookupText_(row[indexes.masterStatus]);
      if (nopol && !MASTER_UNIT_MANAGEMENT.MASTER_STATUSES.includes(status)) {
        status = 'ACTIVE';
        backfilledStatus++;
      }
      unitIds.push([safeSheetText_(unitId)]);
      statuses.push([safeSheetText_(status)]);
      createdAt.push([row[indexes.masterCreatedAt] || '']);
      createdBy.push([safeSheetText_(row[indexes.masterCreatedBy] || (nopol ? migrationActor : ''))]);
      updatedAt.push([row[indexes.masterUpdatedAt] || '']);
      updatedBy.push([safeSheetText_(row[indexes.masterUpdatedBy] || (nopol ? migrationActor : ''))]);
      lastMovement.push([safeSheetText_(row[indexes.lastMovementId] || '')]);
    });
    sheet.getRange(2, indexes.unitId + 1, rowCount, 1).setValues(unitIds);
    sheet.getRange(2, indexes.masterStatus + 1, rowCount, 1).setValues(statuses);
    sheet.getRange(2, indexes.masterCreatedAt + 1, rowCount, 1).setValues(createdAt);
    sheet.getRange(2, indexes.masterCreatedBy + 1, rowCount, 1).setValues(createdBy);
    sheet.getRange(2, indexes.masterUpdatedAt + 1, rowCount, 1).setValues(updatedAt);
    sheet.getRange(2, indexes.masterUpdatedBy + 1, rowCount, 1).setValues(updatedBy);
    sheet.getRange(2, indexes.lastMovementId + 1, rowCount, 1).setValues(lastMovement);
  }

  const dataRows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, indexes.masterStatus + 1, dataRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(MASTER_UNIT_MANAGEMENT.MASTER_STATUSES, true)
      .setAllowInvalid(false).setHelpText('ACTIVE = tampil di monitoring. INACTIVE = disimpan sebagai history master.').build()
  );
  trySetNumberFormat_(sheet.getRange(2, indexes.masterCreatedAt + 1, dataRows, 1), 'dd mmmm yyyy, HH.mm.ss');
  trySetNumberFormat_(sheet.getRange(2, indexes.masterUpdatedAt + 1, dataRows, 1), 'dd mmmm yyyy, HH.mm.ss');
  runNonCritical_('validasi Status Unit Master Unit Management', function () { applyMasterUnitStatusValidation_(sheet, indexes.statusUnit); });
  runNonCritical_('validasi Status 3L Master Unit Management', function () { applyStatus3LValidation_(sheet, indexes); });
  runNonCritical_('hapus proteksi otomatis lama', removeAutomaticSheetProtections_);

  const movement = getOrCreateSheet_(spreadsheet, MASTER_UNIT_MANAGEMENT.MOVEMENT_SHEET);
  formatMasterUnitManagementSheet_(movement, MASTER_UNIT_MANAGEMENT.MOVEMENT_HEADERS, '#0B5D3A');
  const requests = getOrCreateSheet_(spreadsheet, MASTER_UNIT_MANAGEMENT.REQUEST_SHEET);
  formatMasterUnitManagementSheet_(requests, MASTER_UNIT_MANAGEMENT.REQUEST_HEADERS, '#6D28D9');
  return { addedColumns: added, backfilledIds: backfilledIds, backfilledStatus: backfilledStatus };
}

function masterConfigCoverageColumns_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return null;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const aliases = {
    operatingPoint: ['OPERATING POINT', 'OPERATIONAL POINT', 'OP'],
    opc: ['NAMA OPC/WHC', 'OPC/WHC', 'NAMA OPC', 'NAMA WHC'],
    toh: ['NAMA TOH/WOH', 'TOH/WOH', 'NAMA TOH', 'NAMA WOH']
  };
  const findIndex = names => {
    const normalized = names.map(normalizeHeaderKey_);
    return headers.findIndex(header => normalized.includes(normalizeHeaderKey_(header)));
  };
  const indexes = {
    operatingPoint: findIndex(aliases.operatingPoint),
    opc: findIndex(aliases.opc),
    toh: findIndex(aliases.toh)
  };
  return { headers: headers, indexes: indexes };
}

function normalizeCoverageText_(value) {
  return cleanText_(value)
    .replace(/[‐‑‒–—―]/g, '-')
    .replace(/\s+-\s+/g, ' - ')
    .replace(/\s+/g, ' ')
    .toUpperCase();
}

function coverageOperatingPointKey_(value) {
  const text = normalizeCoverageText_(value);
  if (!text) return '';
  const parts = text.split(' - ');
  const candidate = cleanText_(parts[0]).replace(/\s+/g, '');
  // Operating Point pada workbook ini memakai kode stabil di depan nama
  // (contoh KLTDNLU1 - KLATEN DELANGGU 1). Gunakan kode tersebut sebagai
  // primary key agar perubahan spacing/dash/nama display tidak menimbulkan
  // false OP_NOT_IN_COVERAGE. Bila format tidak memiliki kode, fallback ke teks.
  if (parts.length > 1 && /^[A-Z0-9][A-Z0-9._/]{1,39}$/.test(candidate)) {
    return 'CODE:' + candidate;
  }
  return 'TEXT:' + text;
}

function coverageTohKey_(value) {
  return normalizeCoverageText_(value)
    .replace(/[.,]/g, '')
    .replace(/\s+/g, ' ');
}

function masterConfigCoverageDirectory_() {
  if (KM_REQUEST_CONFIG_COVERAGE_) return KM_REQUEST_CONFIG_COVERAGE_;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.CONFIG);
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return [];
  const directory = masterConfigCoverageColumns_(sheet);
  if (!directory || directory.indexes.operatingPoint < 0) return [];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
  const result = [];
  const seen = new Set();
  values.forEach(row => {
    const operatingPoint = cleanText_(row[directory.indexes.operatingPoint]);
    if (!operatingPoint) return;
    const item = {
      operatingPoint: operatingPoint,
      opc: directory.indexes.opc >= 0 ? cleanText_(row[directory.indexes.opc]) : '',
      toh: directory.indexes.toh >= 0 ? cleanText_(row[directory.indexes.toh]) : '',
      source: 'MASTER_CONFIG'
    };
    item.operatingPointKey = coverageOperatingPointKey_(item.operatingPoint);
    item.tohKey = coverageTohKey_(item.toh);
    const key = [item.operatingPointKey, normalizeLookupText_(item.opc), item.tohKey].join('|');
    if (seen.has(key)) return;
    seen.add(key);
    result.push(item);
  });
  KM_REQUEST_CONFIG_COVERAGE_ = result;
  return result;
}

function masterUnitCoverageDirectory_(snapshot) {
  if (KM_REQUEST_UNIT_COVERAGE_) return KM_REQUEST_UNIT_COVERAGE_;
  let source = snapshot;
  if (!source) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
    if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return [];
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    source = { indexes: buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES),
      rows: sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues() };
  }
  const indexes = source.indexes;
  if (indexes.operatingPoint < 0) return [];
  const values = source.rows;
  const result = [];
  const seen = new Set();
  values.forEach(row => {
    const operatingPoint = cleanText_(row[indexes.operatingPoint]);
    if (!operatingPoint) return;
    const toh = indexes.toh >= 0 ? cleanText_(row[indexes.toh]) : '';
    const item = { operatingPoint: operatingPoint, opc: '', toh: toh, source: 'MASTER_UNIT' };
    item.operatingPointKey = coverageOperatingPointKey_(item.operatingPoint);
    item.tohKey = coverageTohKey_(item.toh);
    const key = [item.operatingPointKey, item.tohKey].join('|');
    if (seen.has(key)) return;
    seen.add(key);
    result.push(item);
  });
  KM_REQUEST_UNIT_COVERAGE_ = result;
  return result;
}

function effectiveMasterUnitCoverageDirectory_() {
  if (KM_REQUEST_EFFECTIVE_COVERAGE_) return KM_REQUEST_EFFECTIVE_COVERAGE_;
  const masterRows = masterUnitCoverageDirectory_();
  const configRows = masterConfigCoverageDirectory_();
  const configByPair = new Map();
  configRows.forEach(item => {
    const pairKey = (item.operatingPointKey || coverageOperatingPointKey_(item.operatingPoint)) + '|' +
      (item.tohKey || coverageTohKey_(item.toh));
    if (!configByPair.has(pairKey)) configByPair.set(pairKey, item);
  });
  const result = [];
  const seen = new Set();
  masterRows.forEach(item => {
    const pairKey = (item.operatingPointKey || coverageOperatingPointKey_(item.operatingPoint)) + '|' +
      (item.tohKey || coverageTohKey_(item.toh));
    const config = configByPair.get(pairKey);
    const merged = Object.assign({}, item, {
      opc: (config && config.opc) || item.opc || '',
      source: config ? 'MASTER_UNIT+MASTER_CONFIG' : 'MASTER_UNIT'
    });
    const key = [pairKey, normalizeLookupText_(merged.opc)].join('|');
    if (seen.has(key)) return;
    seen.add(key);
    result.push(merged);
  });
  const representedPairs = new Set(result.map(item => item.operatingPointKey + '|' + item.tohKey));
  configRows.forEach(item => {
    const pairKey = (item.operatingPointKey || coverageOperatingPointKey_(item.operatingPoint)) + '|' +
      (item.tohKey || coverageTohKey_(item.toh));
    if (representedPairs.has(pairKey)) return;
    representedPairs.add(pairKey);
    const fallback = Object.assign({}, item, { source: 'MASTER_CONFIG_FALLBACK' });
    const key = [pairKey, normalizeLookupText_(fallback.opc)].join('|');
    if (seen.has(key)) return;
    seen.add(key);
    result.push(fallback);
  });
  KM_REQUEST_EFFECTIVE_COVERAGE_ = result;
  return result;
}

function resolveMasterUnitCoverage_(operatingPoint, toh) {
  const opKey = coverageOperatingPointKey_(operatingPoint);
  const tohKey = coverageTohKey_(toh);
  const directory = effectiveMasterUnitCoverageDirectory_();
  const candidates = directory.filter(item =>
    (item.operatingPointKey || coverageOperatingPointKey_(item.operatingPoint)) === opKey
  );
  if (!candidates.length) {
    throw new Error('Operating Point belum ditemukan pada MASTER_UNIT maupun referensi Area & Coverage: ' + cleanText_(operatingPoint));
  }
  if (!tohKey) {
    const uniqueToh = uniqueClean_(candidates.map(item => item.toh).filter(Boolean));
    if (uniqueToh.length === 1) return candidates.find(item => cleanText_(item.toh)) || candidates[0];
    throw new Error('TOH/WOH wajib dipilih untuk Operating Point ' + cleanText_(operatingPoint) + '.');
  }
  const match = candidates.find(item =>
    (item.tohKey || coverageTohKey_(item.toh)) === tohKey
  );
  if (match) return match;

  // Izinkan pembentukan pasangan pertama untuk OP baru hanya ketika OP belum
  // mempunyai mapping TOH dan TOH yang dipilih sudah dikenal di directory.
  const hasMappedToh = candidates.some(item => coverageTohKey_(item.toh));
  if (!hasMappedToh) {
    const knownToh = directory.find(item =>
      (item.tohKey || coverageTohKey_(item.toh)) === tohKey && cleanText_(item.toh)
    );
    if (knownToh) {
      return Object.assign({}, candidates[0], {
        operatingPoint: candidates[0].operatingPoint,
        operatingPointKey: opKey,
        toh: knownToh.toh,
        tohKey: knownToh.tohKey || tohKey,
        source: 'NEW_OPERATING_POINT_PAIR'
      });
    }
  }
  throw new Error('TOH/WOH ' + cleanText_(toh) + ' belum terdaftar pada MASTER_UNIT / referensi coverage untuk Operating Point ' + cleanText_(operatingPoint) + '.');
}

function masterUnitRecordSnapshot_(row, indexes) {
  return {
    unitId: displayValue_(row[indexes.unitId]),
    nopol: normalizeDisplayNopol_(row[indexes.nopol]),
    unitType: indexes.unitType >= 0 ? displayValue_(row[indexes.unitType]) : '',
    operatingPoint: indexes.operatingPoint >= 0 ? displayValue_(row[indexes.operatingPoint]) : '',
    toh: indexes.toh >= 0 ? displayValue_(row[indexes.toh]) : '',
    masterStatus: displayValue_(row[indexes.masterStatus]) || 'ACTIVE',
    statusUnit: indexes.statusUnit >= 0 ? displayValue_(row[indexes.statusUnit]) : '',
    kmOdometer: indexes.kmOdometer >= 0 ? parseUnitKm_(row[indexes.kmOdometer]) : null,
    calibrationNumber: indexes.calibrationNumber >= 0 ? parseCalibrationNumber_(row[indexes.calibrationNumber]) : 0,
    status3L: indexes.status3L >= 0 ? displayValue_(row[indexes.status3L]) : '',
    masterCreatedAt: indexes.masterCreatedAt >= 0 ? formatDateTime_(row[indexes.masterCreatedAt]) : '',
    masterCreatedBy: indexes.masterCreatedBy >= 0 ? displayValue_(row[indexes.masterCreatedBy]) : '',
    masterUpdatedAt: indexes.masterUpdatedAt >= 0 ? formatDateTime_(row[indexes.masterUpdatedAt]) : '',
    masterUpdatedBy: indexes.masterUpdatedBy >= 0 ? displayValue_(row[indexes.masterUpdatedBy]) : '',
    lastMovementId: indexes.lastMovementId >= 0 ? displayValue_(row[indexes.lastMovementId]) : ''
  };
}

function findMasterUnitManagementTarget_(sheet, key) {
  if (!sheet || sheet.getLastRow() < 2) throw new Error('MASTER_UNIT belum memiliki data.');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  if (indexes.unitId < 0 || indexes.masterStatus < 0 || indexes.nopol < 0) throw new Error('Master Unit Management belum diinisialisasi. Jalankan Repair Struktur.');
  const target = cleanText_(key);
  const targetId = normalizeLookupText_(target);
  const targetPlate = normalizeNopol_(target);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const matches = [];
  values.forEach((row, offset) => {
    if ((targetId && normalizeLookupText_(row[indexes.unitId]) === targetId) ||
        (targetPlate && normalizeNopol_(row[indexes.nopol]) === targetPlate)) matches.push({ rowNumber: offset + 2, row: row });
  });
  if (matches.length !== 1) throw new Error(matches.length ? 'Unit ditemukan lebih dari satu. Periksa duplicate UNIT_ID/Nopol.' : 'Unit tidak ditemukan.');
  return { headers: headers, indexes: indexes, rowNumber: matches[0].rowNumber, row: matches[0].row, snapshot: masterUnitRecordSnapshot_(matches[0].row, indexes) };
}

function appendMasterUnitMovement_(entry) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, MASTER_UNIT_MANAGEMENT.MOVEMENT_SHEET);
  const map = ensureHistoryHeaders_(sheet, MASTER_UNIT_MANAGEMENT.MOVEMENT_HEADERS);
  const row = new Array(sheet.getLastColumn()).fill('');
  const put = (header, value) => { if (map[header]) row[map[header] - 1] = value == null ? '' : value; };
  Object.keys(entry || {}).forEach(header => put(header, entry[header]));
  ensureSheetCapacity_(sheet, 1);
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  return displayValue_(entry['MOVEMENT ID']);
}

function readMasterUnitMovements_(limit, unitId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_UNIT_MANAGEMENT.MOVEMENT_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const map = ensureHistoryHeaders_(sheet, MASTER_UNIT_MANAGEMENT.MOVEMENT_HEADERS);
  const rowCount = Math.min(Math.max(1, Number(limit) || 200), sheet.getLastRow() - 1);
  const start = Math.max(2, sheet.getLastRow() - rowCount + 1);
  const values = sheet.getRange(start, 1, sheet.getLastRow() - start + 1, sheet.getLastColumn()).getValues().reverse();
  return values.filter(row => !unitId || normalizeLookupText_(row[map['UNIT ID'] - 1]) === normalizeLookupText_(unitId)).map(row => ({
    movementId: displayValue_(row[map['MOVEMENT ID'] - 1]),
    timestamp: formatDateTime_(row[map['TIMESTAMP'] - 1]),
    unitId: displayValue_(row[map['UNIT ID'] - 1]), nopol: displayValue_(row[map['NOMOR POLISI'] - 1]),
    movementType: displayValue_(row[map['MOVEMENT TYPE'] - 1]),
    previousOperatingPoint: displayValue_(row[map['PREVIOUS OPERATING POINT'] - 1]), newOperatingPoint: displayValue_(row[map['NEW OPERATING POINT'] - 1]),
    previousToh: displayValue_(row[map['PREVIOUS TOH/WOH'] - 1]), newToh: displayValue_(row[map['NEW TOH/WOH'] - 1]),
    previousType: displayValue_(row[map['PREVIOUS TYPE'] - 1]), newType: displayValue_(row[map['NEW TYPE'] - 1]),
    effectiveDate: formatDateOnly_(row[map['EFFECTIVE DATE'] - 1]), reason: displayValue_(row[map['REASON'] - 1]), note: displayValue_(row[map['NOTE'] - 1]),
    updatedBy: displayValue_(row[map['UPDATED BY'] - 1]), source: displayValue_(row[map['SOURCE'] - 1]), requestId: displayValue_(row[map['REQUEST ID'] - 1])
  }));
}

function readMasterUnitChangeRequests_(limit) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_UNIT_MANAGEMENT.REQUEST_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const map = ensureHistoryHeaders_(sheet, MASTER_UNIT_MANAGEMENT.REQUEST_HEADERS);
  const count = Math.min(Math.max(1, Number(limit) || 200), sheet.getLastRow() - 1);
  const start = Math.max(2, sheet.getLastRow() - count + 1);
  return sheet.getRange(start, 1, sheet.getLastRow() - start + 1, sheet.getLastColumn()).getValues().reverse().map(row => ({
    requestId: displayValue_(row[map['REQUEST ID'] - 1]), timestamp: formatDateTime_(row[map['TIMESTAMP'] - 1]),
    requestedBy: displayValue_(row[map['REQUESTED BY'] - 1]), requestedRole: displayValue_(row[map['REQUESTED ROLE'] - 1]),
    action: displayValue_(row[map['ACTION'] - 1]), unitId: displayValue_(row[map['UNIT ID'] - 1]), nopol: displayValue_(row[map['NOMOR POLISI'] - 1]),
    proposed: parseJsonObject_(row[map['PROPOSED DATA JSON'] - 1]), effectiveDate: formatDateOnly_(row[map['EFFECTIVE DATE'] - 1]),
    reason: displayValue_(row[map['REASON'] - 1]), note: displayValue_(row[map['NOTE'] - 1]), status: displayValue_(row[map['STATUS'] - 1]),
    reviewedAt: formatDateTime_(row[map['REVIEWED AT'] - 1]), reviewedBy: displayValue_(row[map['REVIEWED BY'] - 1]), reviewNote: displayValue_(row[map['REVIEW NOTE'] - 1]),
    appliedAt: formatDateTime_(row[map['APPLIED AT'] - 1]), appliedBy: displayValue_(row[map['APPLIED BY'] - 1]), movementId: displayValue_(row[map['MOVEMENT ID'] - 1])
  }));
}

function parseJsonObject_(value) {
  if (!value) return {};
  try { const parsed = typeof value === 'string' ? JSON.parse(value) : value; return parsed && typeof parsed === 'object' ? parsed : {}; }
  catch (error) { return {}; }
}

function masterUnitDataQuality_(sheet, indexes) {
  const masterCoverage = masterUnitCoverageDirectory_();
  const configCoverage = masterConfigCoverageDirectory_();
  const effectiveCoverage = effectiveMasterUnitCoverageDirectory_();
  const photoRecoveryIndex = buildOdometerPhotoRecoveryIndex_(SpreadsheetApp.getActiveSpreadsheet());
  const masterOps = new Set(masterCoverage
    .map(item => item.operatingPointKey || coverageOperatingPointKey_(item.operatingPoint))
    .filter(Boolean));
  const masterPairs = new Set(masterCoverage
    .filter(item => cleanText_(item.toh))
    .map(item => (item.operatingPointKey || coverageOperatingPointKey_(item.operatingPoint)) + '|' +
      (item.tohKey || coverageTohKey_(item.toh))));
  const configOps = new Set(configCoverage
    .map(item => item.operatingPointKey || coverageOperatingPointKey_(item.operatingPoint))
    .filter(Boolean));
  const configPairs = new Set(configCoverage
    .filter(item => cleanText_(item.toh))
    .map(item => (item.operatingPointKey || coverageOperatingPointKey_(item.operatingPoint)) + '|' +
      (item.tohKey || coverageTohKey_(item.toh))));

  const issues = [];
  const counts = {};
  const add = (code, severity, rowNumber, nopol, message, meta) => {
    counts[code] = (counts[code] || 0) + 1;
    if (issues.length < 300) {
      issues.push(Object.assign({
        code: code, severity: severity, row: rowNumber,
        nopol: nopol, message: message
      }, meta || {}));
    }
  };

  const coverageDiagnostics = {
    sourceOfTruth: 'MASTER_UNIT',
    masterCoverageRows: masterCoverage.length,
    masterOperatingPoints: masterOps.size,
    masterTohPairs: masterPairs.size,
    configReferenceRows: configCoverage.length,
    configReferenceOperatingPoints: configOps.size,
    configReferenceTohPairs: configPairs.size,
    effectiveCoverageRows: effectiveCoverage.length,
    canonicalKeyMode: 'MASTER_UNIT_PRIMARY_CONFIG_FALLBACK',
    note: 'Mapping existing divalidasi dari MASTER_UNIT. MASTER_CONFIG hanya referensi/fallback untuk pilihan master baru.'
  };

  if (!sheet || sheet.getLastRow() < 2) {
    return {
      score: 100, issueUnits: 0, counts: counts, issues: issues,
      coverageDiagnostics: coverageDiagnostics,
      repairableCounts: { doneNotReady: 0, photoIdNotSynced: 0, photoIntegrity: 0 }
    };
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const nopolMap = {};
  const idMap = {};
  const issueUnitKeys = new Set();
  const missingOpMap = {};
  const missingTohMap = {};
  const totalUnitRows = rows.filter(row => normalizeNopol_(row[indexes.nopol])).length;

  rows.forEach((row, offset) => {
    const rowNumber = offset + 2;
    const nopol = normalizeDisplayNopol_(row[indexes.nopol]);
    const nopolKey = normalizeNopol_(nopol);
    if (!nopolKey) return;

    const unitId = cleanText_(row[indexes.unitId]);
    const masterStatus = normalizeLookupText_(row[indexes.masterStatus]);
    const op = indexes.operatingPoint >= 0 ? cleanText_(row[indexes.operatingPoint]) : '';
    const toh = indexes.toh >= 0 ? cleanText_(row[indexes.toh]) : '';
    const status3L = indexes.status3L >= 0 ? normalizeLookupText_(row[indexes.status3L]) : '';
    const statusUnit = indexes.statusUnit >= 0 ? normalizeLookupText_(row[indexes.statusUnit]) : '';
    const status3LUpdatedBy = indexes.status3LUpdatedBy >= 0
      ? status3LActorDisplay_(displayValue_(row[indexes.status3LUpdatedBy])) : '';
    const calibration = indexes.calibrationNumber >= 0 ? parseCalibrationNumber_(row[indexes.calibrationNumber]) : 0;
    const mark = (code, severity, message, meta) => {
      add(code, severity, rowNumber, nopol, message, meta);
      issueUnitKeys.add(nopolKey);
    };

    if (!unitId) {
      mark('MISSING_UNIT_ID', 'HIGH', 'UNIT ID belum tersedia.');
    } else {
      const idKey = normalizeLookupText_(unitId);
      idMap[idKey] = (idMap[idKey] || []).concat({ row: rowNumber, nopol: nopol, nopolKey: nopolKey });
    }
    nopolMap[nopolKey] = (nopolMap[nopolKey] || []).concat(rowNumber);

    if (!MASTER_UNIT_MANAGEMENT.MASTER_STATUSES.includes(masterStatus)) {
      mark('INVALID_MASTER_STATUS', 'HIGH', 'MASTER STATUS harus ACTIVE / INACTIVE.');
    }
    if (indexes.unitType >= 0 && !cleanText_(row[indexes.unitType])) {
      mark('MISSING_UNIT_TYPE', 'MEDIUM', 'Tipe Unit kosong.');
    }

    // MASTER_UNIT adalah sumber kebenaran coverage existing. Karena itu record tidak
    // lagi dinyatakan mismatch hanya karena MASTER_CONFIG tidak memuat pasangan yang sama.
    if (!op) {
      missingOpMap['MISSING'] = (missingOpMap['MISSING'] || 0) + 1;
      mark('MISSING_OPERATING_POINT', 'HIGH', 'Operating Point kosong pada MASTER_UNIT.');
    }
    if (op && !toh) {
      const opKey = coverageOperatingPointKey_(op) || normalizeLookupText_(op);
      if (!missingTohMap[opKey]) missingTohMap[opKey] = { operatingPoint: op, count: 0 };
      missingTohMap[opKey].count++;
      mark('MISSING_TOH', 'MEDIUM', 'NAMA TOH/WOH kosong pada MASTER_UNIT.');
    }

    if (masterStatus === 'INACTIVE' && calibration && status3L && status3L !== 'DONE') {
      mark('INACTIVE_ACTIVE_CALIBRATION', 'HIGH', 'Unit inactive masih memiliki siklus kalibrasi aktif.');
    }
    if (status3L === 'DONE' && statusUnit !== normalizeLookupText_('Ready to Depart')) {
      mark('DONE_NOT_READY', 'HIGH',
        'Status 3L sudah DONE tetapi Status Unit belum Ready to Depart. Issue ini aman diperbaiki otomatis.',
        { repairHint: 'DONE_READY_INVARIANT' });
    }
    if (status3L === 'DONE' &&
        normalizeLookupText_(status3LUpdatedBy) !== normalizeLookupText_(KM_STATUS_3L_SHEET_FALLBACK)) {
      mark('DONE_ATTRIBUTION_NOT_TEAM_3L', 'HIGH',
        'Status 3L DONE wajib menggunakan attribution operasional Team 3L.',
        { repairHint: 'DONE_TEAM_3L_ATTRIBUTION', currentUpdatedBy: status3LUpdatedBy });
    }

    const rawPhotoId = indexes.kmPhotoId >= 0 ? cleanText_(row[indexes.kmPhotoId]) : '';
    const rawPhotoAddedAt = indexes.kmPhotoAddedAt >= 0 ? row[indexes.kmPhotoAddedAt] : '';
    const rawPhotoAddedBy = indexes.kmPhotoAddedBy >= 0 ? cleanText_(row[indexes.kmPhotoAddedBy]) : '';
    if (!calibration && (rawPhotoId || rawPhotoAddedAt || rawPhotoAddedBy)) {
      mark('PHOTO_WITHOUT_CALIBRATION', 'HIGH',
        'Metadata foto tersedia tetapi KALIBRASI KE kosong; relasi siklus tidak dapat dibuktikan.');
    }

    // Photo Integrity: dashboard membandingkan MASTER_UNIT dengan pointer
    // canonical per siklus, bukan seluruh ID mentah di ledger. Konflik legacy
    // tetap dapat diaudit tetapi tidak menjadi false-positive setelah ada
    // PHOTO_RECONCILIATION append-only.
    if (calibration && indexes.kmPhotoId >= 0) {
      const kmDate = indexes.kmDate >= 0 ? row[indexes.kmDate] : '';
      const masterPhotoId = normalizeOdometerPhotoId_(rawPhotoId);
      const photoSubmitter = indexes.kmUpdatedBy >= 0 ? cleanText_(row[indexes.kmUpdatedBy]) : '';
      const photoResolution = resolveOdometerPhotoRecoveryCandidate_(photoRecoveryIndex, nopol, calibration, kmDate, photoSubmitter);
      const canonicalPhotoId = photoResolution.status === 'RECOVERABLE'
        ? normalizeOdometerPhotoId_(photoResolution.candidate && photoResolution.candidate.photoId) : '';
      const photoKey = analyticsUnitKey_(nopol, calibration);
      const rawPhotoIds = photoKey && photoRecoveryIndex.byCalibration[photoKey]
        ? Object.keys(photoRecoveryIndex.byCalibration[photoKey]) : [];

      if (rawPhotoId && !masterPhotoId) {
        mark('PHOTO_ID_INVALID', 'HIGH',
          'FOTO ODOMETER ID bukan ID Drive atau URL Google Drive yang valid.',
          { repairHint: 'PHOTO_INTEGRITY', rawPhotoId: rawPhotoId });
      } else if (masterPhotoId && rawPhotoId !== masterPhotoId) {
        mark('PHOTO_ID_NOT_NORMALIZED', 'MEDIUM',
          'FOTO ODOMETER ID masih berbentuk URL/teks legacy dan dapat dinormalisasi menjadi ID Drive.',
          { repairHint: 'PHOTO_INTEGRITY', normalizedPhotoId: masterPhotoId });
      }
      if (!masterPhotoId && (rawPhotoAddedAt || rawPhotoAddedBy)) {
        mark('PHOTO_METADATA_ORPHAN', 'HIGH',
          'FOTO DITAMBAHKAN AT / FOTO DIUNGGAH OLEH terisi tanpa FOTO ODOMETER ID yang valid.');
      }

      if (!masterPhotoId) {
        if (photoResolution.status === 'RECOVERABLE') {
          mark('PHOTO_ID_NOT_SYNCED', 'MEDIUM',
            'Bukti foto canonical tersedia di HISTORY_KALIBRASI_KM, tetapi FOTO ODOMETER ID pada MASTER_UNIT kosong. Aman disinkronkan otomatis.',
            { repairHint: 'PHOTO_ID_FROM_HISTORY', photoId: canonicalPhotoId, resolution: photoResolution.match });
        } else if (photoResolution.status === 'AMBIGUOUS') {
          mark('PHOTO_ID_AMBIGUOUS', 'HIGH',
            'History memiliki lebih dari satu kandidat foto dan belum ada pointer canonical. Jalankan Repair Photo Integrity atau audit manual.',
            { repairHint: 'PHOTO_INTEGRITY', historyPhotoIds: (photoResolution.candidates || []).map(item => item.photoId) });
        } else {
          mark('PHOTO_EVIDENCE_MISSING', 'HIGH',
            'FOTO ODOMETER ID kosong dan ledger HISTORY_KALIBRASI_KM tidak memiliki bukti foto untuk siklus ini.');
        }
      } else {
        const metadataCandidate = odometerPhotoCandidateForId_(
          photoRecoveryIndex, nopol, calibration, masterPhotoId, kmDate
        );
        const metadata = compareOdometerPhotoMetadata_(rawPhotoAddedAt, rawPhotoAddedBy, metadataCandidate);
        const metadataMeta = {
          repairHint: 'PHOTO_INTEGRITY',
          photoId: masterPhotoId,
          historyEvidence: Boolean(metadataCandidate)
        };
        if (metadata.missingAt) {
          mark('PHOTO_ADDED_AT_MISSING', 'MEDIUM',
            'FOTO DITAMBAHKAN AT kosong untuk pointer foto aktif.', metadataMeta);
        } else if (metadata.invalidAt) {
          mark('PHOTO_ADDED_AT_INVALID', 'HIGH',
            'FOTO DITAMBAHKAN AT bukan tanggal/waktu yang valid.', metadataMeta);
        } else if (metadata.futureAt) {
          mark('PHOTO_ADDED_AT_FUTURE', 'HIGH',
            'FOTO DITAMBAHKAN AT berada di masa depan.', metadataMeta);
        } else if (metadata.mismatchAt) {
          mark('PHOTO_ADDED_AT_MISMATCH', 'MEDIUM',
            'FOTO DITAMBAHKAN AT tidak sesuai dengan evidence untuk ID foto yang sama di HISTORY_KALIBRASI_KM.', metadataMeta);
        }
        if (metadata.missingBy) {
          mark('PHOTO_UPLOADER_MISSING', 'MEDIUM',
            'FOTO DIUNGGAH OLEH kosong untuk pointer foto aktif.', metadataMeta);
        } else if (metadata.mismatchBy) {
          mark('PHOTO_UPLOADER_MISMATCH', 'MEDIUM',
            'FOTO DIUNGGAH OLEH tidak sesuai dengan evidence untuk ID foto yang sama di HISTORY_KALIBRASI_KM.', metadataMeta);
        }
      }

      if (masterPhotoId && photoResolution.status === 'RECOVERABLE' && canonicalPhotoId && canonicalPhotoId !== masterPhotoId) {
        mark('PHOTO_ID_MISMATCH', 'HIGH',
          'Pointer foto aktif MASTER_UNIT berbeda dengan canonical history. Repair akan mempertahankan pointer aktif dan menulis rekonsiliasi append-only.',
          { repairHint: 'PHOTO_INTEGRITY', masterPhotoId: masterPhotoId, canonicalPhotoId: canonicalPhotoId, historyPhotoIds: rawPhotoIds });
      } else if (masterPhotoId && photoResolution.status === 'AMBIGUOUS') {
        mark('PHOTO_ID_AMBIGUOUS', 'HIGH',
          'History memiliki beberapa ID foto tanpa pointer canonical. MASTER_UNIT tetap dipertahankan sampai rekonsiliasi ditulis.',
          { repairHint: 'PHOTO_INTEGRITY', masterPhotoId: masterPhotoId,
            historyPhotoIds: (photoResolution.candidates || []).map(item => item.photoId) });
      }
    }
  });

  Object.keys(nopolMap).forEach(key => {
    if (nopolMap[key].length > 1) {
      nopolMap[key].forEach(row => {
        add('DUPLICATE_NOPOL', 'HIGH', row, key, 'Nomor polisi duplikat.');
        issueUnitKeys.add(key);
      });
    }
  });
  Object.keys(idMap).forEach(key => {
    if (idMap[key].length <= 1) return;
    idMap[key].forEach(item => {
      add('DUPLICATE_UNIT_ID', 'HIGH', item.row, item.nopol, 'UNIT ID duplikat: ' + key);
      if (item.nopolKey) issueUnitKeys.add(item.nopolKey);
    });
  });

  coverageDiagnostics.missingOperatingPointRows = Number(missingOpMap.MISSING || 0);
  coverageDiagnostics.missingTohGroups = Object.keys(missingTohMap).map(key => missingTohMap[key])
    .sort((a, b) => b.count - a.count || a.operatingPoint.localeCompare(b.operatingPoint, 'id'));

  const total = totalUnitRows;
  const issueUnits = issueUnitKeys.size;
  const score = total ? Math.max(0, Math.round((1 - issueUnits / total) * 1000) / 10) : 100;
  const photoIntegrity = auditOdometerPhotoIntegrity_(sheet, indexes, photoRecoveryIndex);
  return {
    score: score,
    issueUnits: issueUnits,
    counts: counts,
    issues: issues,
    coverageDiagnostics: coverageDiagnostics,
    photoIntegrity: photoIntegrity,
    repairableCounts: {
      doneNotReady: Number(counts.DONE_NOT_READY || 0),
      doneAttributionNotTeam3L: Number(counts.DONE_ATTRIBUTION_NOT_TEAM_3L || 0),
      photoIdNotSynced: Number(counts.PHOTO_ID_NOT_SYNCED || 0),
      photoIntegrity: Number(photoIntegrity.recoverable || 0) + Number(photoIntegrity.mismatch || 0) +
        Number(photoIntegrity.ambiguous || 0) + Number(photoIntegrity.normalizedIdNeeded || 0) +
        Number(photoIntegrity.metadataRepairable || 0)
    }
  };
}

function getAdminMasterUnitManagementData(sessionToken, includeQuality) {
  const actor = requireAnyMasterUnitManagementPermission_(sessionToken);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  if (!sheet) throw new Error('MASTER_UNIT belum tersedia.');
  if (!masterUnitManagementSchemaReady_(sheet)) {
    throw new Error('Master Unit Management belum siap. SUPER_ADMIN perlu menjalankan Repair Struktur satu kali terlebih dahulu.');
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  if (indexes.unitId < 0 || indexes.masterStatus < 0) throw new Error('Master Unit Management belum siap. SUPER_ADMIN perlu menjalankan Repair Struktur.');
  const values = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues() : [];
  const units = [];
  masterUnitCoverageDirectory_({ indexes: indexes, rows: values });
  const stats = { total: 0, active: 0, inactive: 0, calibrated: 0, uncalibrated: 0, movedThisMonth: 0, newThisMonth: 0 };
  const monthKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMM');
  values.forEach(row => {
    const nopol = cleanText_(row[indexes.nopol]);
    if (!nopol) return;
    const snapshot = masterUnitRecordSnapshot_(row, indexes);
    if (!actor.permissions.includes('MANAGE_MASTER_UNIT') && !actor.permissions.includes('APPROVE_MASTER_UNIT_CHANGE') && !canAccessOperatingPoint_(actor, snapshot.operatingPoint)) return;
    units.push(snapshot);
    stats.total++;
    if (normalizeLookupText_(snapshot.masterStatus) === 'INACTIVE') stats.inactive++; else stats.active++;
    if (snapshot.calibrationNumber) stats.calibrated++; else stats.uncalibrated++;
    const created = parseDateTimeInput_(snapshot.masterCreatedAt);
    if (created && Utilities.formatDate(created, Session.getScriptTimeZone(), 'yyyyMM') === monthKey) stats.newThisMonth++;
  });
  let movements = readMasterUnitMovements_(2000);
  if (!actor.permissions.includes('MANAGE_MASTER_UNIT') && !actor.permissions.includes('APPROVE_MASTER_UNIT_CHANGE')) {
    const visibleUnitIds = new Set(units.map(item => normalizeLookupText_(item.unitId)).filter(Boolean));
    movements = movements.filter(item => visibleUnitIds.has(normalizeLookupText_(item.unitId)));
  }
  stats.movedThisMonth = movements.filter(item => item.movementType === 'TRANSFER_OP' && parseDateTimeInput_(item.timestamp) && Utilities.formatDate(parseDateTimeInput_(item.timestamp), Session.getScriptTimeZone(), 'yyyyMM') === monthKey).length;
  const qualityAllowed = actor.permissions.includes('MANAGE_MASTER_UNIT') || actor.permissions.includes('APPROVE_MASTER_UNIT_CHANGE');
  const quality = includeQuality === true && qualityAllowed
    ? masterUnitDataQuality_(sheet, indexes) : { deferred: qualityAllowed, score: null, issueUnits: 0, counts: {}, issues: [] };
  return {
    permissions: {
      manage: actor.permissions.includes('MANAGE_MASTER_UNIT'),
      request: actor.permissions.includes('REQUEST_MASTER_UNIT_CHANGE'),
      approve: actor.permissions.includes('APPROVE_MASTER_UNIT_CHANGE')
    },
    approvalEnabled: masterUnitApprovalEnabled_(),
    stats: stats,
    units: units,
    coverage: effectiveMasterUnitCoverageDirectory_(),
    movements: movements,
    requests: actor.permissions.includes('APPROVE_MASTER_UNIT_CHANGE') ? readMasterUnitChangeRequests_(300) : readMasterUnitChangeRequests_(100).filter(item => normalizeEmail_(item.requestedBy) === normalizeEmail_(actor.email)),
    quality: quality,
    metadataReady: true
  };
}

function validateMasterUnitBasePayload_(payload, action) {
  const data = payload || {};
  const normalizedAction = normalizeLookupText_(action || data.action).replace(/\s+/g, '_');
  const normalizedEffectiveDate = normalizeDateCellServer_(data.effectiveDate);
  if (cleanText_(data.effectiveDate) && !normalizedEffectiveDate) throw new Error('Tanggal efektif tidak valid.');
  const result = {
    action: normalizedAction,
    unitId: cleanText_(data.unitId),
    nopol: normalizeDisplayNopol_(data.nopol),
    unitType: cleanText_(data.unitType),
    operatingPoint: cleanText_(data.operatingPoint),
    toh: cleanText_(data.toh),
    effectiveDate: normalizedEffectiveDate || new Date(),
    reason: cleanText_(data.reason),
    note: cleanText_(data.note)
  };
  if (result.note.length > 1000) throw new Error('Catatan maksimal 1000 karakter.');
  if (result.reason.length > 300) throw new Error('Reason maksimal 300 karakter.');
  if (result.unitType.length > 80) throw new Error('Tipe Unit maksimal 80 karakter.');
  if (normalizedAction === 'ADD') {
    if (!normalizeNopol_(result.nopol)) throw new Error('Nomor polisi wajib dan formatnya tidak valid.');
    if (!result.unitType) throw new Error('Tipe Unit wajib diisi.');
    if (!result.operatingPoint) throw new Error('Operating Point wajib diisi.');
    const coverage = resolveMasterUnitCoverage_(result.operatingPoint, result.toh);
    result.operatingPoint = coverage.operatingPoint;
    result.toh = coverage.toh;
  }
  if (normalizedAction === 'TRANSFER') {
    if (!result.unitId && !result.nopol) throw new Error('UNIT ID / Nopol wajib dipilih untuk transfer.');
    if (!result.operatingPoint) throw new Error('Operating Point tujuan wajib dipilih.');
    if (!result.reason) throw new Error('Reason transfer wajib diisi.');
    const coverage = resolveMasterUnitCoverage_(result.operatingPoint, result.toh);
    result.operatingPoint = coverage.operatingPoint;
    result.toh = coverage.toh;
  }
  if (['DEACTIVATE','REACTIVATE'].includes(normalizedAction) && !result.reason) throw new Error('Reason perubahan status master wajib diisi.');
  return result;
}

function normalizeDateCellServer_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const makeDate = (year, month, day) => {
    const y = Number(year), m = Number(month), d = Number(day);
    if (![y, m, d].every(Number.isInteger) || y < 1900 || y > 9999 || m < 1 || m > 12 || d < 1 || d > 31) return null;
    const date = new Date(y, m - 1, d);
    return date.getFullYear() === y && date.getMonth() === m - 1 && date.getDate() === d ? date : null;
  };
  const text = cleanText_(value).replace(/\u00a0/g, ' ');
  if (!text) return null;
  const numericText = text.replace(',', '.');
  if (/^\d+(?:[.,]\d+)?$/.test(text) && Number(numericText) > 0 && Number(numericText) < 100000) {
    // Excel 1900 date system. Gunakan komponen UTC dari serial agar tanggal
    // kalender tidak bergeser mengikuti timezone runtime Apps Script.
    const wall = new Date(Date.UTC(1899, 11, 30) + Math.floor(Number(numericText)) * 86400000);
    return makeDate(wall.getUTCFullYear(), wall.getUTCMonth() + 1, wall.getUTCDate());
  }
  let match = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T].*)?$/);
  if (match) return makeDate(match[1], match[2], match[3]);
  match = text.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})(?:[ T].*)?$/);
  if (match) {
    const shortYear = Number(match[3]);
    const year = match[3].length === 2 ? (shortYear <= 69 ? 2000 + shortYear : 1900 + shortYear) : shortYear;
    // Locale aplikasi Indonesia: format teks ambigu diperlakukan DD/MM/YYYY.
    return makeDate(year, match[2], match[1]);
  }
  match = text.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (match) return makeDate(match[1], match[2], match[3]);
  const parsed = parseDateTimeInput_(text);
  return parsed instanceof Date && !isNaN(parsed.getTime()) ? parsed : null;
}

function writeMasterUnitMetadata_(sheet, target, actor, movementId, isCreate) {
  const now = new Date();
  const indexes = target.indexes;
  if (isCreate) {
    sheet.getRange(target.rowNumber, indexes.masterCreatedAt + 1).setValue(now);
    sheet.getRange(target.rowNumber, indexes.masterCreatedBy + 1).setValue(safeSheetText_(actor.email));
  }
  sheet.getRange(target.rowNumber, indexes.masterUpdatedAt + 1).setValue(now);
  sheet.getRange(target.rowNumber, indexes.masterUpdatedBy + 1).setValue(safeSheetText_(actor.email));
  if (movementId) sheet.getRange(target.rowNumber, indexes.lastMovementId + 1).setValue(safeSheetText_(movementId));
}

function applyMasterUnitManagementAction_(action, payload, actor, options) {
  const settings = options || {};
  const data = validateMasterUnitBasePayload_(payload, action);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  if (!masterUnitManagementSchemaReady_(sheet)) {
    throw new Error('Master Unit Management belum siap. Jalankan Repair Struktur sebagai SUPER_ADMIN sebelum melakukan perubahan.');
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  const now = new Date();
  const requestId = cleanText_(settings.requestId);
  const source = settings.source || 'ADMIN_DASHBOARD';

  if (data.action === 'ADD') {
    const plateKey = normalizeNopol_(data.nopol);
    if (sheet.getLastRow() > 1) {
      const existing = sheet.getRange(2, indexes.nopol + 1, sheet.getLastRow() - 1, 1).getDisplayValues();
      const found = existing.findIndex(row => normalizeNopol_(row[0]) === plateKey);
      if (found >= 0) {
        const existingStatus = normalizeLookupText_(sheet.getRange(found + 2, indexes.masterStatus + 1).getDisplayValue());
        throw new Error(existingStatus === 'INACTIVE' ? 'Unit sudah ada dalam status INACTIVE. Gunakan Reactivate, jangan membuat duplicate.' : 'Nomor polisi sudah terdaftar di MASTER_UNIT.');
      }
    }
    const row = new Array(sheet.getLastColumn()).fill('');
    const unitId = newMasterUnitId_();
    row[indexes.unitId] = unitId;
    row[indexes.masterStatus] = 'ACTIVE';
    row[indexes.nopol] = safeSheetText_(data.nopol);
    if (indexes.unitType >= 0) row[indexes.unitType] = safeSheetText_(data.unitType);
    row[indexes.operatingPoint] = safeSheetText_(data.operatingPoint);
    row[indexes.toh] = safeSheetText_(data.toh);
    // v24.1: unit baru belum masuk workflow operasional/kalibrasi.
    // STATUS UNIT sengaja dibiarkan kosong sampai user mengubah status atau membuat pengajuan Kalibrasi KM.
    if (indexes.statusUnit >= 0) row[indexes.statusUnit] = '';
    row[indexes.masterCreatedAt] = now; row[indexes.masterCreatedBy] = safeSheetText_(actor.email);
    row[indexes.masterUpdatedAt] = now; row[indexes.masterUpdatedBy] = safeSheetText_(actor.email);
    const movementId = newMasterUnitMovementId_(now);
    row[indexes.lastMovementId] = movementId;
    ensureSheetCapacity_(sheet, 1);
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    appendMasterUnitMovement_({
      'MOVEMENT ID': movementId, 'TIMESTAMP': now, 'UNIT ID': unitId, 'NOMOR POLISI': data.nopol, 'MOVEMENT TYPE': 'CREATED',
      'NEW OPERATING POINT': data.operatingPoint, 'NEW TOH/WOH': data.toh, 'NEW TYPE': data.unitType, 'NEW NOMOR POLISI': data.nopol,
      'EFFECTIVE DATE': data.effectiveDate, 'REASON': data.reason || 'Penambahan master unit', 'NOTE': data.note,
      'UPDATED BY': safeSheetText_(actor.email), 'SOURCE': source, 'REQUEST ID': requestId
    });
    return { action: 'ADD', unitId: unitId, nopol: data.nopol, movementId: movementId, initialStatusUnit: '' };
  }

  const target = findMasterUnitManagementTarget_(sheet, data.unitId || data.nopol);
  const before = target.snapshot;
  if (!actor.permissions.includes('MANAGE_MASTER_UNIT') && !actor.permissions.includes('APPROVE_MASTER_UNIT_CHANGE') && !canAccessOperatingPoint_(actor, before.operatingPoint)) {
    throw new Error('ACCESS_DENIED: unit berada di luar scope Operating Point akun Anda.');
  }
  const movementId = newMasterUnitMovementId_(now);
  if (data.action === 'EDIT') {
    if (data.unitType) sheet.getRange(target.rowNumber, indexes.unitType + 1).setValue(safeSheetText_(data.unitType));
    writeMasterUnitMetadata_(sheet, target, actor, movementId, false);
    appendMasterUnitMovement_({
      'MOVEMENT ID': movementId, 'TIMESTAMP': now, 'UNIT ID': before.unitId, 'NOMOR POLISI': before.nopol, 'MOVEMENT TYPE': 'EDIT_MASTER',
      'PREVIOUS OPERATING POINT': before.operatingPoint, 'NEW OPERATING POINT': before.operatingPoint,
      'PREVIOUS TOH/WOH': before.toh, 'NEW TOH/WOH': before.toh, 'PREVIOUS TYPE': before.unitType, 'NEW TYPE': data.unitType || before.unitType,
      'PREVIOUS NOMOR POLISI': before.nopol, 'NEW NOMOR POLISI': before.nopol, 'EFFECTIVE DATE': data.effectiveDate,
      'REASON': data.reason || 'Edit master unit', 'NOTE': data.note, 'UPDATED BY': safeSheetText_(actor.email), 'SOURCE': source, 'REQUEST ID': requestId
    });
  } else if (data.action === 'TRANSFER') {
    if (normalizeLookupText_(before.masterStatus) !== 'ACTIVE') throw new Error('Unit INACTIVE harus direactivate sebelum transfer.');
    if (normalizeLookupText_(before.operatingPoint) === normalizeLookupText_(data.operatingPoint) && normalizeLookupText_(before.toh) === normalizeLookupText_(data.toh)) {
      throw new Error('Operating Point / TOH tujuan sama dengan coverage saat ini.');
    }
    sheet.getRange(target.rowNumber, indexes.operatingPoint + 1).setValue(safeSheetText_(data.operatingPoint));
    sheet.getRange(target.rowNumber, indexes.toh + 1).setValue(safeSheetText_(data.toh));
    writeMasterUnitMetadata_(sheet, target, actor, movementId, false);
    appendMasterUnitMovement_({
      'MOVEMENT ID': movementId, 'TIMESTAMP': now, 'UNIT ID': before.unitId, 'NOMOR POLISI': before.nopol, 'MOVEMENT TYPE': 'TRANSFER_OP',
      'PREVIOUS OPERATING POINT': before.operatingPoint, 'NEW OPERATING POINT': data.operatingPoint,
      'PREVIOUS TOH/WOH': before.toh, 'NEW TOH/WOH': data.toh, 'PREVIOUS TYPE': before.unitType, 'NEW TYPE': before.unitType,
      'PREVIOUS NOMOR POLISI': before.nopol, 'NEW NOMOR POLISI': before.nopol, 'EFFECTIVE DATE': data.effectiveDate,
      'REASON': data.reason, 'NOTE': data.note, 'UPDATED BY': safeSheetText_(actor.email), 'SOURCE': source, 'REQUEST ID': requestId
    });
  } else if (['DEACTIVATE','REACTIVATE'].includes(data.action)) {
    const nextStatus = data.action === 'DEACTIVATE' ? 'INACTIVE' : 'ACTIVE';
    if (normalizeLookupText_(before.masterStatus) === nextStatus) throw new Error('MASTER STATUS unit sudah ' + nextStatus + '.');
    if (nextStatus === 'INACTIVE' && before.calibrationNumber && normalizeLookupText_(before.status3L) !== 'DONE') {
      throw new Error('Unit masih memiliki kalibrasi aktif. Selesaikan / finalisasi kalibrasi sebelum Deactivate.');
    }
    sheet.getRange(target.rowNumber, indexes.masterStatus + 1).setValue(nextStatus);
    writeMasterUnitMetadata_(sheet, target, actor, movementId, false);
    appendMasterUnitMovement_({
      'MOVEMENT ID': movementId, 'TIMESTAMP': now, 'UNIT ID': before.unitId, 'NOMOR POLISI': before.nopol,
      'MOVEMENT TYPE': nextStatus === 'ACTIVE' ? 'REACTIVATED' : 'DEACTIVATED',
      'PREVIOUS OPERATING POINT': before.operatingPoint, 'NEW OPERATING POINT': before.operatingPoint,
      'PREVIOUS TOH/WOH': before.toh, 'NEW TOH/WOH': before.toh, 'PREVIOUS TYPE': before.unitType, 'NEW TYPE': before.unitType,
      'PREVIOUS NOMOR POLISI': before.nopol, 'NEW NOMOR POLISI': before.nopol, 'EFFECTIVE DATE': data.effectiveDate,
      'REASON': data.reason, 'NOTE': data.note, 'UPDATED BY': safeSheetText_(actor.email), 'SOURCE': source, 'REQUEST ID': requestId
    });
  } else {
    throw new Error('Action Master Unit tidak didukung: ' + data.action);
  }
  SpreadsheetApp.flush();
  return { action: data.action, unitId: before.unitId, nopol: before.nopol, movementId: movementId };
}

function saveAdminMasterUnitRecord(payload, sessionToken, deferData) {
  const actor = requirePermission_('MANAGE_MASTER_UNIT', '', { token: sessionToken });
  const action = normalizeLookupText_(payload && payload.action || 'EDIT').replace(/\s+/g, '_');
  if (!['ADD','EDIT'].includes(action)) throw new Error('Gunakan endpoint khusus untuk action ' + action + '.');
  const lock = acquireScriptLock_(30000, 'Master Unit ' + action);
  let result;
  try {
    result = applyMasterUnitManagementAction_(action, payload, actor);
    logSystemActivity_(result.nopol, '', 'MASTER_UNIT_' + action, 'MASTER_UNIT', '', result.unitId,
      'Master Unit Management: ' + action, actor, true);
    invalidateAppCaches_();
  } finally { lock.releaseLock(); }
  return { success: true, result: result, dataDeferred: deferData === true,
    data: deferData === true ? null : getAdminMasterUnitManagementData(sessionToken) };
}

function transferAdminMasterUnit(payload, sessionToken, deferData) {
  const actor = requirePermission_('MANAGE_MASTER_UNIT', '', { token: sessionToken });
  const lock = acquireScriptLock_(30000, 'Transfer Master Unit');
  let result;
  try {
    result = applyMasterUnitManagementAction_('TRANSFER', payload, actor);
    logSystemActivity_(result.nopol, '', 'MASTER_UNIT_TRANSFERRED', 'OPERATING POINT', '', cleanText_(payload.operatingPoint),
      'Transfer Operating Point. ' + cleanText_(payload.reason), actor, true);
    invalidateAppCaches_();
  } finally { lock.releaseLock(); }
  return { success: true, result: result, dataDeferred: deferData === true,
    data: deferData === true ? null : getAdminMasterUnitManagementData(sessionToken) };
}

function setAdminMasterUnitActive(payload, sessionToken, deferData) {
  const actor = requirePermission_('MANAGE_MASTER_UNIT', '', { token: sessionToken });
  const action = payload && payload.active === false ? 'DEACTIVATE' : 'REACTIVATE';
  const lock = acquireScriptLock_(30000, action + ' Master Unit');
  let result;
  try {
    result = applyMasterUnitManagementAction_(action, payload, actor);
    logSystemActivity_(result.nopol, '', 'MASTER_UNIT_' + action, 'MASTER STATUS', action === 'DEACTIVATE' ? 'ACTIVE' : 'INACTIVE', action === 'DEACTIVATE' ? 'INACTIVE' : 'ACTIVE',
      cleanText_(payload.reason), actor, true);
    invalidateAppCaches_();
  } finally { lock.releaseLock(); }
  return { success: true, result: result, dataDeferred: deferData === true,
    data: deferData === true ? null : getAdminMasterUnitManagementData(sessionToken) };
}

function setAdminMasterUnitApprovalEnabled(enabled, sessionToken) {
  const actor = requirePermission_('APPROVE_MASTER_UNIT_CHANGE', '', { token: sessionToken });
  const lock = acquireScriptLock_(20000, 'Master Unit Approval Configuration');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet_(ss, KM_APP.SHEETS.CONFIG);
    if (sheet.getMaxColumns() < 6) sheet.insertColumnsAfter(sheet.getMaxColumns(), 6 - sheet.getMaxColumns());
    sheet.getRange('E1:F1').setValues([['PARAMETER','VALUE']]);
    const last = Math.max(5, sheet.getLastRow());
    const values = sheet.getRange(2, 5, Math.max(4, last - 1), 2).getDisplayValues();
    let rowNumber = 0;
    values.forEach((row, offset) => { if (normalizeLookupText_(row[0]) === MASTER_UNIT_MANAGEMENT.APPROVAL_PARAMETER) rowNumber = offset + 2; });
    if (!rowNumber) {
      const emptyIndex = values.findIndex(row => !cleanText_(row[0]));
      rowNumber = emptyIndex >= 0 ? emptyIndex + 2 : last + 1;
    }
    sheet.getRange(rowNumber, 5, 1, 2).setValues([[MASTER_UNIT_MANAGEMENT.APPROVAL_PARAMETER, enabled ? 'TRUE' : 'FALSE']]);
    invalidateAppCaches_();
    logSystemActivity_('', '', 'MASTER_UNIT_APPROVAL_CONFIG', MASTER_UNIT_MANAGEMENT.APPROVAL_PARAMETER, '', enabled ? 'TRUE' : 'FALSE', 'Optional Approval Workflow diperbarui.', actor, true);
    return { success: true, approvalEnabled: Boolean(enabled) };
  } finally { lock.releaseLock(); }
}

function submitAdminMasterUnitChangeRequest(payload, sessionToken) {
  const actor = requirePermission_('REQUEST_MASTER_UNIT_CHANGE', '', { token: sessionToken });
  if (!masterUnitApprovalEnabled_()) throw new Error('Optional Approval Workflow sedang OFF. Hubungi Admin Bitrans untuk perubahan langsung.');
  const data = validateMasterUnitBasePayload_(payload, payload && payload.action);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  if (!masterUnitManagementSchemaReady_(master)) {
    throw new Error('Master Unit Management belum siap. SUPER_ADMIN perlu menjalankan Repair Struktur sebelum request dapat dibuat.');
  }
  let current = {};
  if (data.action === 'ADD' && master.getLastRow() > 1) {
    const headers = master.getRange(1, 1, 1, master.getLastColumn()).getDisplayValues()[0];
    const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
    const plates = master.getRange(2, indexes.nopol + 1, master.getLastRow() - 1, 1).getDisplayValues();
    if (plates.some(row => normalizeNopol_(row[0]) === normalizeNopol_(data.nopol))) {
      throw new Error('Nomor polisi sudah terdaftar. Gunakan Edit / Transfer / Reactivate sesuai kondisi unit.');
    }
  }
  if (data.action !== 'ADD') {
    const target = findMasterUnitManagementTarget_(master, data.unitId || data.nopol);
    current = target.snapshot;
    if (!canAccessOperatingPoint_(actor, current.operatingPoint) && !actor.permissions.includes('MANAGE_MASTER_UNIT')) {
      throw new Error('ACCESS_DENIED: hanya unit pada scope Operating Point Anda yang dapat diajukan.');
    }
  } else if (!canAccessOperatingPoint_(actor, data.operatingPoint) && !actor.globalScope) {
    throw new Error('ACCESS_DENIED: Operating Point tujuan di luar scope akun Anda.');
  }
  const sheet = getOrCreateSheet_(ss, MASTER_UNIT_MANAGEMENT.REQUEST_SHEET);
  const map = ensureHistoryHeaders_(sheet, MASTER_UNIT_MANAGEMENT.REQUEST_HEADERS);
  const requestId = newMasterUnitRequestId_(new Date());
  const row = new Array(sheet.getLastColumn()).fill('');
  const put = (header, value) => { if (map[header]) row[map[header] - 1] = value == null ? '' : value; };
  put('REQUEST ID', requestId); put('TIMESTAMP', new Date()); put('REQUESTED BY', safeSheetText_(actor.email)); put('REQUESTED ROLE', actor.roles.join(', '));
  put('ACTION', data.action); put('UNIT ID', data.unitId || current.unitId || ''); put('NOMOR POLISI', data.nopol || current.nopol || '');
  put('CURRENT DATA JSON', JSON.stringify(current)); put('PROPOSED DATA JSON', JSON.stringify(data)); put('EFFECTIVE DATE', data.effectiveDate);
  put('REASON', safeSheetText_(data.reason)); put('NOTE', safeSheetText_(data.note)); put('STATUS', 'PENDING');
  ensureSheetCapacity_(sheet, 1); sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  logSystemActivity_(data.nopol || current.nopol || '', '', 'MASTER_UNIT_CHANGE_REQUESTED', data.action, '', requestId, data.reason, actor, true);
  return { success: true, requestId: requestId };
}

function reviewAdminMasterUnitChangeRequest(requestId, decision, reviewNote, sessionToken) {
  const actor = requirePermission_('APPROVE_MASTER_UNIT_CHANGE', '', { token: sessionToken });
  const verdict = normalizeLookupText_(decision);
  if (!['APPROVE','REJECT'].includes(verdict)) throw new Error('Decision harus APPROVE atau REJECT.');
  const lock = acquireScriptLock_(30000, 'Review Master Unit Request');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(MASTER_UNIT_MANAGEMENT.REQUEST_SHEET);
    if (!sheet || sheet.getLastRow() < 2) throw new Error('Request tidak ditemukan.');
    const map = ensureHistoryHeaders_(sheet, MASTER_UNIT_MANAGEMENT.REQUEST_HEADERS);
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const index = values.findIndex(row => normalizeLookupText_(row[map['REQUEST ID'] - 1]) === normalizeLookupText_(requestId));
    if (index < 0) throw new Error('Request ID tidak ditemukan.');
    const rowNumber = index + 2; const row = values[index];
    const status = normalizeLookupText_(row[map['STATUS'] - 1]);
    if (status !== 'PENDING') throw new Error('Request sudah ' + status + ' dan tidak dapat diproses ulang.');
    const now = new Date();
    sheet.getRange(rowNumber, map['REVIEWED AT']).setValue(now);
    sheet.getRange(rowNumber, map['REVIEWED BY']).setValue(safeSheetText_(actor.email));
    sheet.getRange(rowNumber, map['REVIEW NOTE']).setValue(safeSheetText_(cleanText_(reviewNote)));
    if (verdict === 'REJECT') {
      sheet.getRange(rowNumber, map['STATUS']).setValue('REJECTED');
      logSystemActivity_(displayValue_(row[map['NOMOR POLISI'] - 1]), '', 'MASTER_UNIT_REQUEST_REJECTED', displayValue_(row[map['ACTION'] - 1]), requestId, 'REJECTED', cleanText_(reviewNote), actor, true);
      return { success: true, status: 'REJECTED', data: getAdminMasterUnitManagementData(sessionToken) };
    }
    const proposed = parseJsonObject_(row[map['PROPOSED DATA JSON'] - 1]);
    proposed.unitId = proposed.unitId || displayValue_(row[map['UNIT ID'] - 1]);
    proposed.nopol = proposed.nopol || displayValue_(row[map['NOMOR POLISI'] - 1]);
    const action = normalizeLookupText_(row[map['ACTION'] - 1]).replace(/\s+/g, '_');
    if (action !== 'ADD') {
      const currentStored = parseJsonObject_(row[map['CURRENT DATA JSON'] - 1]);
      const master = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
      const currentNow = findMasterUnitManagementTarget_(master, proposed.unitId || proposed.nopol).snapshot;
      const compareFields = ['unitId','nopol','unitType','operatingPoint','toh','masterStatus'];
      const stale = compareFields.some(field => normalizeLookupText_(currentStored[field]) !== normalizeLookupText_(currentNow[field]));
      if (stale) throw new Error('REQUEST_STALE: master unit berubah sejak request dibuat. Reject request ini lalu submit ulang berdasarkan data terbaru.');
    }
    const result = applyMasterUnitManagementAction_(action, proposed, actor, { requestId: requestId, source: 'APPROVED_REQUEST' });
    sheet.getRange(rowNumber, map['STATUS']).setValue('APPROVED');
    sheet.getRange(rowNumber, map['APPLIED AT']).setValue(now);
    sheet.getRange(rowNumber, map['APPLIED BY']).setValue(safeSheetText_(actor.email));
    sheet.getRange(rowNumber, map['MOVEMENT ID']).setValue(safeSheetText_(result.movementId));
    logSystemActivity_(result.nopol, '', 'MASTER_UNIT_REQUEST_APPROVED', result.action, requestId, result.movementId, cleanText_(reviewNote), actor, true);
    invalidateAppCaches_();
    return { success: true, status: 'APPROVED', result: result, data: getAdminMasterUnitManagementData(sessionToken) };
  } finally { lock.releaseLock(); }
}

function parseAdminMasterUnitBulk_(csvText, expectedAction) {
  if (String(csvText || '').length > 5 * 1024 * 1024) throw new Error('Payload Bulk Master Unit maksimal 5 MB.');
  const rows = parseAdminCsv_(csvText);
  if (!rows.length || !rows[0].length) throw new Error('File bulk kosong atau header tidak ditemukan.');
  if (rows.length - 1 > MASTER_UNIT_MANAGEMENT.BULK_MAX_ROWS) throw new Error('Bulk Master Unit maksimal ' + MASTER_UNIT_MANAGEMENT.BULK_MAX_ROWS + ' baris.');
  const headers = rows[0].map(normalizeHeaderKey_);
  const col = names => {
    const candidates = Array.isArray(names) ? names : [names];
    for (let i = 0; i < candidates.length; i++) {
      const index = headers.indexOf(normalizeHeaderKey_(candidates[i]));
      if (index >= 0) return index;
    }
    return -1;
  };
  const normalizedExpectedAction = normalizeLookupText_(expectedAction).replace(/\s+/g, '_');
  if (!['ADD','EDIT','TRANSFER'].includes(normalizedExpectedAction)) throw new Error('Jenis bulk harus ADD, EDIT, atau TRANSFER.');
  const columns = {
    action: col('ACTION'), unitId: col(['UNIT ID','UNIT_ID']), nopol: col(['NOMOR POLISI','NOPOL']),
    unitType: col(['TIPE BARU','TYPE BARU','TIPE','TYPE']),
    operatingPoint: col(['OPERATING POINT TUJUAN','OPERATING POINT']),
    toh: col(['TOH/WOH TUJUAN','TOH / WOH TUJUAN','TOH/WOH','TOH / WOH']),
    effectiveDate: col(['EFFECTIVE DATE','TANGGAL EFEKTIF']), reason: col(['REASON','ALASAN']), note: col(['NOTE','CATATAN'])
  };
  const requiredColumns = normalizedExpectedAction === 'ADD'
    ? ['nopol','unitType','operatingPoint','toh']
    : normalizedExpectedAction === 'EDIT' ? ['unitType'] : ['operatingPoint','toh','effectiveDate','reason'];
  const missingColumns = requiredColumns.filter(name => columns[name] < 0);
  if (normalizedExpectedAction !== 'ADD' && columns.unitId < 0 && columns.nopol < 0) missingColumns.push('UNIT ID / NOMOR POLISI');
  const columnLabels = { nopol: 'NOMOR POLISI', unitType: 'TIPE', operatingPoint: 'OPERATING POINT', toh: 'TOH/WOH', effectiveDate: 'EFFECTIVE DATE', reason: 'REASON' };
  if (missingColumns.length) throw new Error('Template Bulk ' + normalizedExpectedAction + ' tidak sesuai. Kolom belum tersedia: ' + missingColumns.map(name => columnLabels[name] || name).join(', ') + '.');
  const items = rows.slice(1).map((row, offset) => ({ row: row, sourceRow: offset + 2 }))
    .filter(item => item.row.some(cell => cleanText_(cell))).map(item => {
      const row = item.row;
      return {
        sourceRow: item.sourceRow,
        action: columns.action >= 0 ? cleanText_(row[columns.action]) : normalizedExpectedAction,
        nopol: columns.nopol >= 0 ? cleanText_(row[columns.nopol]) : '',
        unitId: columns.unitId >= 0 ? cleanText_(row[columns.unitId]) : '',
        unitType: columns.unitType >= 0 ? cleanText_(row[columns.unitType]) : '',
        operatingPoint: columns.operatingPoint >= 0 ? cleanText_(row[columns.operatingPoint]) : '',
        toh: columns.toh >= 0 ? cleanText_(row[columns.toh]) : '',
        effectiveDate: columns.effectiveDate >= 0 ? cleanText_(row[columns.effectiveDate]) : '',
        reason: columns.reason >= 0 ? cleanText_(row[columns.reason]) : '',
        note: columns.note >= 0 ? cleanText_(row[columns.note]) : '',
        expectedAction: normalizedExpectedAction
      };
    });
  if (!items.length) throw new Error('Sheet DATA belum berisi baris data.');
  return items;
}

function validateAdminMasterUnitBulkRows_(items) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  if (!masterUnitManagementSchemaReady_(sheet)) throw new Error('Master Unit Management belum siap. Jalankan Repair Struktur terlebih dahulu.');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  const values = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues() : [];
  const directoryByPlate = {};
  const directoryById = {};
  const duplicatePlates = Object.create(null);
  const duplicateIds = Object.create(null);
  values.forEach(row => {
    const snapshot = masterUnitRecordSnapshot_(row, indexes);
    const nopol = normalizeNopol_(snapshot.nopol);
    const unitId = normalizeLookupText_(snapshot.unitId);
    if (nopol && directoryByPlate[nopol]) duplicatePlates[nopol] = true;
    if (unitId && directoryById[unitId]) duplicateIds[unitId] = true;
    if (nopol) directoryByPlate[nopol] = snapshot;
    if (unitId) directoryById[unitId] = snapshot;
  });
  const touchedUnits = Object.create(null);
  return items.map(item => {
    const action = normalizeLookupText_(item.action).replace(/\s+/g, '_');
    const errors = [];
    const expectedAction = normalizeLookupText_(item.expectedAction).replace(/\s+/g, '_');
    const beforeById = item.unitId ? directoryById[normalizeLookupText_(item.unitId)] || null : null;
    const beforeByPlate = item.nopol ? directoryByPlate[normalizeNopol_(item.nopol)] || null : null;
    let before = beforeById || beforeByPlate || null;
    let after = null;
    let payload = item;
    try {
      if (action !== expectedAction) throw new Error('ACTION pada baris tidak sesuai dengan jenis Bulk ' + expectedAction + '.');
      if (!['ADD','EDIT','TRANSFER'].includes(action)) throw new Error('Jenis bulk hanya mendukung ADD, EDIT, atau TRANSFER.');
      if (item.unitId && duplicateIds[normalizeLookupText_(item.unitId)]) throw new Error('UNIT ID ditemukan lebih dari satu pada MASTER_UNIT. Perbaiki duplicate terlebih dahulu.');
      if (item.nopol && duplicatePlates[normalizeNopol_(item.nopol)]) throw new Error('Nomor polisi ditemukan lebih dari satu pada MASTER_UNIT. Perbaiki duplicate terlebih dahulu.');
      if (item.unitId && !beforeById) throw new Error('UNIT ID tidak ditemukan.');
      if (item.nopol && action !== 'ADD' && !beforeByPlate) throw new Error('Nomor polisi tidak ditemukan.');
      if (beforeById && beforeByPlate && normalizeLookupText_(beforeById.unitId) !== normalizeLookupText_(beforeByPlate.unitId)) throw new Error('UNIT ID dan Nomor Polisi menunjuk unit yang berbeda.');
      if (action !== 'ADD' && !item.unitId && !item.nopol) throw new Error('Isi minimal salah satu: UNIT ID atau Nomor Polisi.');
      if (action === 'EDIT' && !item.unitType) throw new Error('Tipe Baru wajib diisi.');
      if (action === 'TRANSFER' && !item.effectiveDate) throw new Error('Effective Date transfer wajib diisi.');
      const data = validateMasterUnitBasePayload_(Object.assign({}, item, { action: action }), action);
      if (action === 'ADD' && before) throw new Error('Nopol sudah terdaftar.');
      if (action !== 'ADD' && !before) throw new Error('Nopol belum terdaftar.');
      if (action === 'ADD') after = Object.assign({}, data, { masterStatus: 'ACTIVE' });
      if (action === 'EDIT') {
        if (normalizeLookupText_(before.unitType) === normalizeLookupText_(data.unitType)) throw new Error('Tipe Baru sama dengan tipe unit saat ini.');
        after = Object.assign({}, before, { unitType: data.unitType });
      }
      if (action === 'TRANSFER') after = Object.assign({}, before, { operatingPoint: data.operatingPoint, toh: data.toh });
      if (action === 'TRANSFER' && before && normalizeLookupText_(before.masterStatus) !== 'ACTIVE') throw new Error('Unit INACTIVE tidak dapat ditransfer.');
      if (action === 'TRANSFER' && normalizeLookupText_(before.operatingPoint) === normalizeLookupText_(data.operatingPoint) &&
          normalizeLookupText_(before.toh) === normalizeLookupText_(data.toh)) throw new Error('Operating Point / TOH tujuan sama dengan coverage saat ini.');
      const touchKey = action === 'ADD' ? normalizeNopol_(data.nopol) : normalizeLookupText_(before.unitId);
      if (touchedUnits[touchKey]) throw new Error('Unit yang sama sudah digunakan pada row ' + touchedUnits[touchKey] + '. Satu unit hanya boleh muncul sekali per file.');
      touchedUnits[touchKey] = item.sourceRow;
      payload = Object.assign({}, item, data, {
        action: action,
        unitId: before ? before.unitId : data.unitId,
        nopol: before ? before.nopol : data.nopol
      });
      if (action === 'ADD') {
        directoryByPlate[normalizeNopol_(data.nopol)] = after;
      } else {
        directoryByPlate[normalizeNopol_(before.nopol)] = after;
        directoryById[normalizeLookupText_(before.unitId)] = after;
      }
    } catch (error) { errors.push(String(error && error.message ? error.message : error)); }
    return { sourceRow: item.sourceRow, action: action, unitId: item.unitId, nopol: item.nopol, before: before, after: after, errors: errors, valid: errors.length === 0, payload: payload };
  });
}

function previewAdminMasterUnitBulk(csvText, expectedAction, sessionToken) {
  requirePermission_('MANAGE_MASTER_UNIT', '', { token: sessionToken });
  const items = parseAdminMasterUnitBulk_(csvText, expectedAction);
  const rows = validateAdminMasterUnitBulkRows_(items);
  // payload internal berisi Date untuk EFFECTIVE DATE. Date tidak boleh ikut
  // dikembalikan melalui google.script.run karena respons dapat gagal
  // diserialisasi dan browser akhirnya menerima hasil kosong. Preview hanya
  // membutuhkan snapshot tampilan; payload lengkap divalidasi ulang saat Apply.
  const publicRows = rows.slice(0, 200).map(item => ({
    sourceRow: item.sourceRow,
    action: item.action,
    unitId: item.unitId,
    nopol: item.nopol,
    before: item.before ? {
      unitId: item.before.unitId,
      nopol: item.before.nopol,
      unitType: item.before.unitType,
      operatingPoint: item.before.operatingPoint,
      toh: item.before.toh,
      masterStatus: item.before.masterStatus
    } : null,
    after: item.after ? {
      unitId: item.after.unitId,
      nopol: item.after.nopol,
      unitType: item.after.unitType,
      operatingPoint: item.after.operatingPoint,
      toh: item.after.toh,
      masterStatus: item.after.masterStatus
    } : null,
    errors: item.errors,
    valid: item.valid
  }));
  return {
    action: normalizeLookupText_(expectedAction),
    rowCount: rows.length,
    validCount: rows.filter(item => item.valid).length,
    errorCount: rows.filter(item => !item.valid).length,
    rows: publicRows,
    valid: rows.length > 0 && rows.every(item => item.valid)
  };
}

function applyAdminMasterUnitBulk(csvText, expectedAction, sessionToken) {
  const actor = requirePermission_('MANAGE_MASTER_UNIT', '', { token: sessionToken });
  const action = normalizeLookupText_(expectedAction).replace(/\s+/g, '_');
  const items = parseAdminMasterUnitBulk_(csvText, action);
  const lock = acquireScriptLock_(60000, 'Bulk Master Unit Management');
  const applied = [];
  try {
    // Validasi ulang dilakukan setelah lock didapat agar preview final dan apply
    // membaca snapshot master yang sama.
    const preview = validateAdminMasterUnitBulkRows_(items);
    const invalid = preview.filter(item => !item.valid);
    if (invalid.length) throw new Error('Bulk dibatalkan. Terdapat ' + invalid.length + ' baris invalid. Perbaiki seluruh error sebelum Apply.');
    preview.forEach(item => applied.push(applyMasterUnitManagementAction_(item.action, item.payload, actor, { source: 'BULK_ADMIN' })));
    SpreadsheetApp.flush();
    invalidateAppCaches_();
    logSystemActivity_('', '', 'MASTER_UNIT_BULK_' + action + '_APPLIED', 'MASTER_UNIT', '', String(applied.length), 'Bulk Master Unit ' + action + ' selesai.', actor, true);
    return { success: true, action: action, applied: applied.length, data: getAdminMasterUnitManagementData(sessionToken) };
  } finally { lock.releaseLock(); }
}


function adminCsvEscape_(value) {
  const text = value == null ? '' : String(value);
  return /[",\n\r]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text;
}

function adminCsvPayload_(filename, headers, rows) {
  const csv = '\ufeff' + [headers].concat(rows).map(row => row.map(adminCsvEscape_).join(',')).join('\r\n');
  return {
    filename: filename,
    mimeType: 'text/csv;charset=utf-8',
    base64: Utilities.base64Encode(Utilities.newBlob(csv, 'text/csv', filename).getBytes()),
    rowCount: rows.length
  };
}

function exportAdministrationDataset(dataset, sessionToken) {
  const actor = requireAdminCenterPermission_('ADMIN_SETUP', sessionToken);
  const key = normalizeLookupText_(dataset).replace(/\s+/g, '_');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  let payload;

  if (key === 'MASTER_UNIT') {
    const sheet = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
    if (!sheet) throw new Error('MASTER_UNIT belum tersedia.');
    const values = sheet.getDataRange().getDisplayValues();
    payload = adminCsvPayload_('MASTER_UNIT_' + stamp + '.csv', values[0] || [], values.slice(1));
  } else if (key === 'AREA_COVERAGE') {
    const data = getAdminAreaCoverageData(sessionToken);
    payload = adminCsvPayload_('AREA_COVERAGE_' + stamp + '.csv',
      ['OPERATING POINT', 'NAMA OPC/WHC', 'NAMA TOH/WOH'],
      data.rows.map(item => [item.operatingPoint, item.opc, item.toh]));
  } else if (key === 'WORKFLOW') {
    const config = getAdminWorkflowConfig(sessionToken);
    const schedule = config.submissionSchedule || {};
    const rows = config.status3LOptions.map(value => ['STATUS_3L', value, ''])
      .concat([
        ['PARAMETER', 'APP_NAME', config.appName],
        ['PARAMETER', 'PAGE_SIZE', config.pageSize],
        ['PARAMETER', 'SCHEMA_VERSION', config.schemaVersion],
        ['PARAMETER', KM_CALIBRATION_SUBMISSION_SCHEDULE.ENABLED, schedule.enabled ? 'TRUE' : 'FALSE'],
        ['PARAMETER', KM_CALIBRATION_SUBMISSION_SCHEDULE.START_AT, schedule.startAt || ''],
        ['PARAMETER', KM_CALIBRATION_SUBMISSION_SCHEDULE.END_AT, schedule.endAt || ''],
        ['PARAMETER', KM_CALIBRATION_SUBMISSION_SCHEDULE.DAILY_START, schedule.dailyStart || ''],
        ['PARAMETER', KM_CALIBRATION_SUBMISSION_SCHEDULE.DAILY_END, schedule.dailyEnd || ''],
        ['PARAMETER', KM_CALIBRATION_SUBMISSION_SCHEDULE.ALLOWED_DAYS,
          (schedule.allowedDays || []).join(',')]
      ]);
    payload = adminCsvPayload_('WORKFLOW_CONFIG_' + stamp + '.csv', ['TYPE', 'KEY', 'VALUE'], rows);
  } else if (key === 'USERS') {
    requirePermission_('MANAGE_USERS', '', { token: sessionToken });
    const users = readUserDirectory_(true).map(publicUserRecord_);
    const headers = ['EMAIL','USERNAME','DISPLAY NAME','ROLES','OPERATING POINTS','ACTIVE','MUST CHANGE PASSWORD','LOCKED UNTIL','LAST LOGIN AT','PASSWORD CHANGED AT'];
    payload = adminCsvPayload_('USERS_SAFE_' + stamp + '.csv', headers, users.map(user => [
      user.email, user.username, user.displayName, (user.roles || []).join('; '),
      (user.operatingPoints || []).join('; '), user.active, user.mustChangePassword,
      user.lockedUntil, user.lastLoginAt, user.passwordChangedAt
    ]));
  } else if (key === 'ROLES') {
    requirePermission_('MANAGE_USERS', '', { token: sessionToken });
    const roles = readRoleDirectory_(true);
    payload = adminCsvPayload_('ROLES_' + stamp + '.csv',
      ['ROLE','DISPLAY NAME','PERMISSIONS','ACTIVE','SYSTEM ROLE'],
      roles.map(role => [role.role, role.displayName, (role.permissions || []).join('; '), role.active, role.systemRole]));
  } else if (key === 'AUDIT_TRAIL') {
    requirePermission_('VIEW_AUDIT', '', { token: sessionToken });
    const rows = getAuditTrail(500, sessionToken);
    payload = adminCsvPayload_('AUDIT_TRAIL_' + stamp + '.csv',
      ['TIMESTAMP','USER EMAIL','USER ROLE','SOURCE','ACTION','FIELD NAME','NOMOR POLISI','OLD VALUE','NEW VALUE','NOTE'],
      rows.map(item => [item.timestamp,item.email,item.role,item.source,item.action,item.fieldName,item.nopol,item.oldValue,item.newValue,item.note]));
  } else if (key === 'ACTIVE_SESSIONS') {
    requirePermission_('MANAGE_USERS', '', { token: sessionToken });
    const sessions = getAdminActiveSessions(sessionToken).sessions;
    payload = adminCsvPayload_('ACTIVE_SESSIONS_' + stamp + '.csv',
      ['SESSION ID','USER','EMAIL','CREATED AT','LAST ACTIVITY AT','EXPIRES AT','STATUS','CURRENT SESSION'],
      sessions.map(item => [item.sessionId,item.displayName,item.email,item.createdAt,item.lastActivityAt,item.expiresAt,item.status,item.current]));
  } else if (key === 'MASTER_UNIT_MOVEMENT') {
    requirePermission_('MANAGE_MASTER_UNIT', '', { token: sessionToken });
    const movementSheet = ss.getSheetByName(MASTER_UNIT_MANAGEMENT.MOVEMENT_SHEET);
    const values = movementSheet && movementSheet.getLastRow() ? movementSheet.getDataRange().getDisplayValues() : [MASTER_UNIT_MANAGEMENT.MOVEMENT_HEADERS];
    payload = adminCsvPayload_('MASTER_UNIT_MOVEMENT_' + stamp + '.csv', values[0] || MASTER_UNIT_MANAGEMENT.MOVEMENT_HEADERS, values.slice(1));
  } else if (key === 'MASTER_UNIT_REQUESTS') {
    requirePermission_('APPROVE_MASTER_UNIT_CHANGE', '', { token: sessionToken });
    const requestSheet = ss.getSheetByName(MASTER_UNIT_MANAGEMENT.REQUEST_SHEET);
    const values = requestSheet && requestSheet.getLastRow() ? requestSheet.getDataRange().getDisplayValues() : [MASTER_UNIT_MANAGEMENT.REQUEST_HEADERS];
    payload = adminCsvPayload_('MASTER_UNIT_REQUESTS_' + stamp + '.csv', values[0] || MASTER_UNIT_MANAGEMENT.REQUEST_HEADERS, values.slice(1));
  } else {
    throw new Error('Dataset export tidak didukung.');
  }

  logSystemActivity_('', '', 'ADMIN_EXPORT', key, '', String(payload.rowCount),
    'Export Administration: ' + key, actor);
  return payload;
}

function parseAdminCsv_(csvText) {
  const text = String(csvText || '');
  if (!text.trim()) throw new Error('File CSV kosong.');
  if (Utilities.newBlob(text).getBytes().length > KM_ADMIN_IMPORT_MAX_BYTES) {
    throw new Error('Ukuran CSV maksimal 2 MB.');
  }
  const rows = Utilities.parseCsv(text.replace(/^\ufeff/, ''));
  if (!rows.length) throw new Error('CSV tidak memiliki data.');
  if (rows.length - 1 > KM_ADMIN_IMPORT_MAX_ROWS) throw new Error('CSV maksimal ' + KM_ADMIN_IMPORT_MAX_ROWS + ' baris data.');
  return rows;
}

function adminCsvHeaderMap_(headers) {
  const map = {};
  (headers || []).forEach((header, index) => {
    const key = normalizeHeaderKey_(header);
    if (key && !Object.prototype.hasOwnProperty.call(map, key)) map[key] = index;
  });
  return map;
}

function previewAdministrationImport(dataset, csvText, sessionToken) {
  requireAdminCenterPermission_('ADMIN_SETUP', sessionToken);
  const key = normalizeLookupText_(dataset).replace(/\s+/g, '_');
  const rows = parseAdminCsv_(csvText);
  const headers = rows[0].map(cleanText_);
  const map = adminCsvHeaderMap_(headers);
  const errors = [];
  if (key === 'AREA_COVERAGE') {
    ['OPERATINGPOINT'].forEach(required => { if (!Object.prototype.hasOwnProperty.call(map, required)) errors.push('Header OPERATING POINT wajib ada.'); });
  } else if (key === 'STATUS_3L_OPTIONS') {
    if (!Object.prototype.hasOwnProperty.call(map, 'STATUSPROSES3L')) errors.push('Header STATUS PROSES 3L wajib ada.');
  } else if (key === 'MASTER_UNIT_DIRECTORY') {
    errors.push('Import MASTER_UNIT_DIRECTORY legacy dinonaktifkan untuk menjaga Movement History. Gunakan Administration > Master Unit > Bulk Management.');
  } else {
    errors.push('Jenis import tidak didukung.');
  }
  return {
    dataset: key,
    headers: headers,
    rowCount: Math.max(0, rows.length - 1),
    sample: rows.slice(1, 6),
    errors: errors,
    valid: errors.length === 0
  };
}

function importMasterUnitDirectory_(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  if (!sheet) throw new Error('MASTER_UNIT belum tersedia.');
  const masterHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(masterHeaders, MASTER_UNIT_FIELD_ALIASES);
  if (indexes.nopol < 0) throw new Error('Header NOMOR POLISI tidak ditemukan di MASTER_UNIT.');

  const csvHeaders = rows[0].map(cleanText_);
  const fieldByCsvColumn = {};
  const allowedFields = ['nopol','operatingPoint','toh','unitType'];
  csvHeaders.forEach((header, column) => {
    const key = normalizeHeaderKey_(header);
    const field = allowedFields.find(name => (MASTER_UNIT_FIELD_ALIASES[name] || []).some(alias => normalizeHeaderKey_(alias) === key));
    if (field && indexes[field] >= 0) fieldByCsvColumn[column] = field;
  });
  if (!Object.values(fieldByCsvColumn).includes('nopol')) throw new Error('Kolom NOMOR POLISI tidak dapat dipetakan ke MASTER_UNIT.');

  const existingRows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues()
    : [];
  const rowByNopol = {};
  existingRows.forEach((row, offset) => {
    const key = normalizeNopol_(row[indexes.nopol]);
    if (!key) return;
    if (rowByNopol[key]) throw new Error('MASTER_UNIT memiliki NOMOR POLISI duplikat: ' + row[indexes.nopol]);
    rowByNopol[key] = offset + 2;
  });

  const newRows = [];
  let inserted = 0;
  let updated = 0;
  rows.slice(1).forEach((csvRow, rowOffset) => {
    const mapped = {};
    Object.keys(fieldByCsvColumn).forEach(column => {
      mapped[fieldByCsvColumn[column]] = cleanText_(csvRow[Number(column)]);
    });
    if (!mapped.nopol) return;
    const key = normalizeNopol_(mapped.nopol);
    if (!key) throw new Error('Nomor polisi invalid pada baris CSV ' + (rowOffset + 2) + '.');
    const targetRow = rowByNopol[key];
    if (targetRow) {
      Object.keys(mapped).forEach(field => {
        if (field === 'nopol' || indexes[field] < 0) return;
        sheet.getRange(targetRow, indexes[field] + 1).setValue(safeSheetText_(mapped[field]));
      });
      updated++;
    } else {
      const row = new Array(sheet.getLastColumn()).fill('');
      Object.keys(mapped).forEach(field => {
        if (indexes[field] >= 0) row[indexes[field]] = safeSheetText_(mapped[field]);
      });
      newRows.push(row);
      rowByNopol[key] = -1;
      inserted++;
    }
  });
  if (newRows.length) {
    ensureSheetCapacity_(sheet, newRows.length);
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, sheet.getLastColumn()).setValues(newRows);
  }
  return { inserted: inserted, updated: updated };
}

function applyAdministrationImport(dataset, csvText, sessionToken) {
  const actor = requireAdminCenterPermission_('ADMIN_SETUP', sessionToken);
  const preview = previewAdministrationImport(dataset, csvText, sessionToken);
  if (!preview.valid) throw new Error(preview.errors.join(' '));
  const key = preview.dataset;
  const rows = parseAdminCsv_(csvText);
  const map = adminCsvHeaderMap_(rows[0]);
  let result;

  if (key === 'AREA_COVERAGE') {
    const payload = rows.slice(1).map(row => ({
      operatingPoint: row[map.OPERATINGPOINT] || '',
      opc: Object.prototype.hasOwnProperty.call(map, 'NAMAOPCWHC') ? row[map.NAMAOPCWHC] : '',
      toh: Object.prototype.hasOwnProperty.call(map, 'NAMATOHWOH') ? row[map.NAMATOHWOH] : ''
    }));
    result = saveAdminAreaCoverage(payload, sessionToken);
  } else if (key === 'STATUS_3L_OPTIONS') {
    const statuses = rows.slice(1).map(row => cleanText_(row[map.STATUSPROSES3L])).filter(Boolean);
    const current = getAdminWorkflowConfig(sessionToken);
    result = saveAdminWorkflowConfig({ appName: current.appName, pageSize: current.pageSize, status3LOptions: statuses }, sessionToken);
  } else if (key === 'MASTER_UNIT_DIRECTORY') {
    throw new Error('Import MASTER_UNIT_DIRECTORY legacy dinonaktifkan. Gunakan Master Unit > Bulk Management agar setiap perubahan tercatat pada Movement History.');
  } else {
    throw new Error('Jenis import tidak didukung.');
  }
  logSystemActivity_('', '', 'ADMIN_IMPORT', key, '', JSON.stringify(result || {}),
    'Import CSV dari Administration.', actor);
  return { success: true, dataset: key, result: result };
}

function getAdminActiveSessions(sessionToken) {
  const actor = requireAdminCenterPermission_('MANAGE_USERS', sessionToken);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.SESSIONS);
  if (!sheet || sheet.getLastRow() < 2) return { sessions: [], stats: { active: 0, expired: 0 } };
  const map = getHeaderMap_(sheet);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const users = readUserDirectory_(true);
  const userById = {};
  users.forEach(user => { userById[user.userId] = user; });
  const now = Date.now();
  let expired = 0;
  const sessions = values.map((row, offset) => {
    const sessionId = cleanText_(getRowValue_(row, map, 'SESSION ID'));
    const userId = cleanText_(getRowValue_(row, map, 'USER ID'));
    const user = userById[userId] || {};
    const createdAtRaw = getRowValue_(row, map, 'CREATED AT');
    const createdAtDate = createdAtRaw instanceof Date ? createdAtRaw : parseDateTimeInput_(createdAtRaw);
    const expiresAtRaw = getRowValue_(row, map, 'EXPIRES AT');
    const expiresAtDate = expiresAtRaw instanceof Date ? expiresAtRaw : parseDateTimeInput_(expiresAtRaw);
    const status = normalizeLookupText_(getRowValue_(row, map, 'STATUS')) || 'UNKNOWN';
    const isExpired = Boolean(sessionId) && status === 'ACTIVE' && expiresAtDate instanceof Date && expiresAtDate.getTime() <= now;
    if (isExpired) expired++;
    return {
      rowNumber: offset + 2,
      sessionId: sessionId,
      userId: userId,
      email: user.email || '',
      displayName: user.displayName || user.username || user.email || userId,
      createdAt: formatDateTime_(createdAtRaw),
      createdAtMs: createdAtDate instanceof Date && !isNaN(createdAtDate.getTime()) ? createdAtDate.getTime() : 0,
      expiresAt: formatDateTime_(expiresAtRaw),
      lastActivityAt: formatDateTime_(getRowValue_(row, map, 'LAST ACTIVITY AT')),
      status: isExpired ? 'EXPIRED' : status,
      current: sessionId === actor.sessionId
    };
  }).filter(item => item.sessionId && (item.status === 'ACTIVE' || item.status === 'EXPIRED'))
    .sort((left, right) => right.createdAtMs - left.createdAtMs)
    .map(item => {
      delete item.createdAtMs;
      return item;
    });
  return {
    sessions: sessions,
    stats: {
      active: sessions.filter(item => item.status === 'ACTIVE').length,
      expired: expired,
      users: uniqueClean_(sessions.filter(item => item.status === 'ACTIVE').map(item => item.email || item.userId)).length
    }
  };
}

function revokeAdminSession(sessionId, sessionToken) {
  const actor = requireAdminCenterPermission_('MANAGE_USERS', sessionToken);
  const target = cleanText_(sessionId);
  if (!target) throw new Error('Session ID wajib diisi.');
  if (target === actor.sessionId) throw new Error('Sesi aktif Anda sendiri harus diakhiri melalui tombol Logout.');
  const record = readSessionById_(target);
  if (!record) throw new Error('Session tidak ditemukan.');
  revokeSessionByRecord_(record, 'REVOKED_BY_ADMIN_CENTER');
  logSystemActivity_('', '', 'ADMIN_SESSION_REVOKED', 'KM_SESSIONS', target, 'REVOKED',
    'Session dicabut dari Administration.', actor);
  return { success: true, sessionId: target };
}

function cleanupExpiredAdminSessions(sessionToken) {
  const actor = requireAdminCenterPermission_('MANAGE_USERS', sessionToken);
  const lock = acquireScriptLock_(30000, 'cleanup sesi kedaluwarsa');
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.SESSIONS);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, cleaned: 0 };
    const map = getHeaderMap_(sheet);
    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    const values = range.getValues();
    const now = Date.now();
    let cleaned = 0;
    values.forEach(row => {
      if (normalizeLookupText_(getRowValue_(row, map, 'STATUS')) !== 'ACTIVE') return;
      const expiresAt = getRowValue_(row, map, 'EXPIRES AT');
      const date = expiresAt instanceof Date ? expiresAt : parseDateTimeInput_(expiresAt);
      if (!(date instanceof Date) || date.getTime() > now) return;
      setRowValue_(row, map, 'STATUS', 'REVOKED');
      setRowValue_(row, map, 'REVOKED AT', new Date());
      setRowValue_(row, map, 'REVOKE REASON', 'EXPIRED_CLEANUP');
      const id = cleanText_(getRowValue_(row, map, 'SESSION ID'));
      if (id) markSessionRevokedInCache_(id);
      cleaned++;
    });
    if (cleaned) range.setValues(values);
    logSystemActivity_('', '', 'ADMIN_SESSION_CLEANUP', 'KM_SESSIONS', '', String(cleaned),
      'Expired session cleanup.', actor, true);
    return { success: true, cleaned: cleaned };
  } finally {
    lock.releaseLock();
  }
}

const ADMIN_HEALTH_CACHE_KEY = 'admin-health-deep-v24';
const ADMIN_HEALTH_CACHE_SECONDS = 300;

function adminDatabaseSheetStats_(ss, existingInventory) {
  const inventory = existingInventory || auditDatabaseSheetInventory_(ss);
  const inventoryByName = inventory.reduce((map, item) => {
    map[item.name] = item;
    return map;
  }, {});
  const sheets = ss.getSheets().map(sheet => {
    const name = sheet.getName();
    const known = inventoryByName[name] || {};
    return {
      name: name,
      rows: Number(known.rows || 0),
      columns: Number(known.columns || 0),
      maxRows: sheet.getMaxRows(),
      maxColumns: sheet.getMaxColumns(),
      hidden: Boolean(known.hidden)
    };
  });
  return { inventory: inventory, sheets: sheets };
}

function adminDatabaseTriggerHealth_(ss) {
  const allTriggers = ScriptApp.getProjectTriggers();
  const editTriggers = allTriggers.filter(trigger =>
    isKalibrasiKmEditTriggerForSpreadsheet_(trigger, ss.getId())
  );
  const watchTriggers = allTriggers.filter(trigger =>
    trigger.getHandlerFunction() === KM_STATUS_3L_WATCH.HANDLER &&
    trigger.getEventType() === ScriptApp.EventType.CLOCK
  );
  const properties = PropertiesService.getScriptProperties();
  const editCanonicalId = properties.getProperty(KM_SHEET_EDIT_TRIGGER.ID_PROPERTY);
  const editEnabled = properties.getProperty(KM_SHEET_EDIT_TRIGGER.ENABLED_PROPERTY) === 'TRUE';
  const watchCanonicalId = properties.getProperty(KM_STATUS_3L_WATCH.ID_PROPERTY);
  const watchEnabled = properties.getProperty(KM_STATUS_3L_WATCH.ENABLED_PROPERTY) === 'TRUE';
  const editReady = editEnabled && editTriggers.length === 1 &&
    (!editCanonicalId || String(editTriggers[0].getUniqueId()) === String(editCanonicalId));
  const watchReady = watchEnabled && watchTriggers.length === 1 &&
    (!watchCanonicalId || String(watchTriggers[0].getUniqueId()) === String(watchCanonicalId));
  const watchLastRunAt = properties.getProperty(KM_STATUS_3L_WATCH.LAST_RUN_PROPERTY) || '';
  const watchLastSuccessAt = properties.getProperty(KM_STATUS_3L_WATCH.LAST_SUCCESS_PROPERTY) || '';
  const watchLastError = properties.getProperty(KM_STATUS_3L_WATCH.LAST_ERROR_PROPERTY) || '';
  const watchErrorCount = Math.max(0, Number(properties.getProperty(KM_STATUS_3L_WATCH.ERROR_COUNT_PROPERTY) || 0));
  const heartbeatDate = parseDateTimeInput_(watchLastSuccessAt || watchLastRunAt);
  const heartbeatAgeMinutes = heartbeatDate instanceof Date && !isNaN(heartbeatDate.getTime())
    ? Math.max(0, Math.round((Date.now() - heartbeatDate.getTime()) / 60000)) : null;
  const heartbeatReady = heartbeatAgeMinutes === null || heartbeatAgeMinutes <= Math.max(5, KM_STATUS_3L_WATCH.INTERVAL_MINUTES * 4);
  return {
    triggerReady: editReady && watchReady && heartbeatReady && watchErrorCount < 3,
    triggerCount: editTriggers.length + watchTriggers.length,
    triggerEnabled: editEnabled && watchEnabled,
    canonicalTriggerIdConfigured: Boolean(editCanonicalId) && Boolean(watchCanonicalId),
    editTriggerReady: editReady,
    editTriggerCount: editTriggers.length,
    watchTriggerReady: watchReady,
    watchTriggerCount: watchTriggers.length,
    watcherBaselineReady: Boolean(properties.getProperty(KM_STATUS_3L_WATCH.LAST_SIGNATURE_PROPERTY)),
    triggerLastFiredAt: properties.getProperty(KM_SHEET_EDIT_TRIGGER.LAST_FIRED_PROPERTY) || '',
    watchLastRunAt: watchLastRunAt,
    watchLastSuccessAt: watchLastSuccessAt,
    watchHeartbeatReady: heartbeatReady,
    watchHeartbeatAgeMinutes: heartbeatAgeMinutes,
    watchErrorCount: watchErrorCount,
    watchLastError: watchLastError
  };
}

function buildAdminDatabaseHealthPayload_(inspection, ss, options) {
  const settings = options || {};
  const stats = adminDatabaseSheetStats_(ss, inspection.inventory);
  const trigger = adminDatabaseTriggerHealth_(ss);
  const inventory = inspection.inventory || stats.inventory || [];
  const historyContext = inspection.historyContext || {};
  const repairSnapshotIntegrity = inspection.repairSnapshotIntegrity || {};
  const analyticsLedger = inspection.analyticsLedger || {};
  const deepAudit = settings.deepAudit === true;
  const ready = deepAudit
    ? Boolean(inspection.ready) && trigger.triggerReady && Boolean(historyContext.ready) &&
      Boolean(repairSnapshotIntegrity.ready) &&
      !inventory.some(item => item.classification === 'LEGACY_MERGE' || item.classification === 'LEGACY_ARCHIVE')
    : null;
  // Avoid sending the three large audit objects twice. They remain available at
  // top level while inspection keeps schema and MASTER_UNIT diagnostics.
  const inspectionPayload = deepAudit ? {
    ready: inspection.ready,
    spreadsheetName: inspection.spreadsheetName,
    timezone: inspection.timezone,
    authPepperConfigured: inspection.authPepperConfigured,
    schemaVersion: inspection.schemaVersion,
    sheets: inspection.sheets || [],
    masterUnit: inspection.masterUnit || {},
    recommendation: inspection.recommendation || ''
  } : inspection;
  return {
    ready: ready,
    auditPending: !deepAudit,
    cached: Boolean(settings.cached),
    inspection: inspectionPayload,
    historyContext: historyContext,
    repairSnapshotIntegrity: repairSnapshotIntegrity,
    analyticsLedger: analyticsLedger,
    inventory: inventory,
    triggerReady: trigger.triggerReady,
    triggerCount: trigger.triggerCount,
    triggerEnabled: trigger.triggerEnabled,
    canonicalTriggerIdConfigured: trigger.canonicalTriggerIdConfigured,
    triggerLastFiredAt: trigger.triggerLastFiredAt,
    editTriggerReady: trigger.editTriggerReady,
    editTriggerCount: trigger.editTriggerCount,
    watchTriggerReady: trigger.watchTriggerReady,
    watchTriggerCount: trigger.watchTriggerCount,
    watcherBaselineReady: trigger.watcherBaselineReady,
    watchLastRunAt: trigger.watchLastRunAt,
    sheets: stats.sheets,
    checkedAt: new Date().toISOString()
  };
}

/** Fast first paint: only metadata/header checks. Full row audits are loaded separately. */
function getAdminDatabaseHealth(sessionToken) {
  requireAdminCenterPermission_('ADMIN_SETUP', sessionToken);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const structure = inspectDatabaseStructure_(ss);
  return buildAdminDatabaseHealthPayload_(structure, ss, { deepAudit: false });
}

/** Expensive audit, cached for repeat visits. forceRefresh bypasses the five-minute cache. */
function getAdminDatabaseHealthAudit(sessionToken, forceRefresh) {
  requireAdminCenterPermission_('ADMIN_SETUP', sessionToken);
  if (!forceRefresh) {
    const cached = getCachedJson_(ADMIN_HEALTH_CACHE_KEY);
    if (cached) {
      cached.cached = true;
      return cached;
    }
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inspection = inspectDatabaseKalibrasiKm_();
  const result = buildAdminDatabaseHealthPayload_(inspection, ss, { deepAudit: true });
  putCachedJson_(ADMIN_HEALTH_CACHE_KEY, result, ADMIN_HEALTH_CACHE_SECONDS);
  return result;
}

function runAdminStandardizeFoundation_(actor) {
  const lock = acquireScriptLock_(30000, 'standardisasi foundation database');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const headerOrder = standardizeEnterpriseHeaderOrder_(ss);
    const formats = formatDatabaseTables_(ss);
    const accessMode = removeAutomaticSheetProtections_();
    SpreadsheetApp.flush();
    return { success: true, headerOrder: headerOrder, formats: formats, accessMode: accessMode };
  } finally { lock.releaseLock(); }
}

/** Repair ringan untuk trigger dan validation tanpa mengubah data bisnis/history. */
function runAdminRepairTriggerAccess_(actor) {
  const lock = acquireScriptLock_(30000, 'repair trigger dan access');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const master = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
    let validation = { statusUnit: false, status3L: false };
    if (master && master.getLastColumn()) {
      const headers = master.getRange(1, 1, 1, master.getLastColumn()).getDisplayValues()[0];
      const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
      validation.statusUnit = indexes.statusUnit >= 0 &&
        applyMasterUnitStatusValidation_(master, indexes.statusUnit);
      if (indexes.status3L >= 0) {
        applyStatus3LValidation_(master, indexes);
        validation.status3L = true;
      }
    }
    invalidateSecurityCaches_();
    KM_REQUEST_ROLE_DIRECTORY_ = null;
    KM_REQUEST_USER_DIRECTORY_ = null;
    KM_REQUEST_ACTIVE_ROLE_NAMES_ = null;
    const editTrigger = ensureKalibrasiKmEditTrigger_();
    const status3LWatchTrigger = ensureStatus3LWatchTrigger_();
    SpreadsheetApp.flush();
    return {
      success: true,
      safeTriggerOnly: true,
      existingDataChanged: false,
      editTrigger: editTrigger,
      status3LWatchTrigger: status3LWatchTrigger,
      validation: validation,
      health: adminDatabaseTriggerHealth_(ss)
    };
  } finally { lock.releaseLock(); }
}

function runAdminRepairDataIntegrity_(actor, options) {
  const settings = options || {};
  const backup = settings.skipBackup ? null : createAdministrationBackup_(actor, 'BEFORE_DATA_INTEGRITY_REPAIR');
  const photo = repairCalibrationPhotoIntegrity_(actor);
  const history = repairCalibrationHistoryContextIntegrity_(actor, { skipBackup: true });
  const analytics = rebuildAnalyticsCalibrationLedgerCore_(actor);
  SpreadsheetApp.flush();
  return { success: true, backup: backup, photoIntegrity: photo, historyContext: history, analyticsLedger: analytics };
}

function runAdminOptimizeStructure_(actor, options) {
  const settings = options || {};
  const backup = settings.skipBackup ? null : createAdministrationBackup_(actor, 'BEFORE_STRUCTURE_OPTIMIZATION');
  const legacy = consolidateLegacyDatabaseSheets_(actor, { skipBackup: true });
  const clean = cleanDatabaseStructure_(actor, { skipBackup: true, skipFoundation: true });
  const compact = compactDatabaseGrid_();
  const foundation = runAdminStandardizeFoundation_(actor);
  SpreadsheetApp.flush();
  return { success: true, backup: backup, legacy: legacy, clean: clean, compact: compact, foundation: foundation };
}

function runAdminFullSafeRepair_(actor) {
  const backup = createAdministrationBackup_(actor, 'BEFORE_FULL_SAFE_REPAIR');
  // Schema first. setupDatabaseKalibrasiKMCore_ is evidence-first in v22 and
  // no longer mirrors MASTER_UNIT into history as proof.
  const schema = setupDatabaseKalibrasiKMCore_(actor);
  const workflow = repairStatus3LWorkflowDataIntegrityUnsafe_(actor);
  const dataIntegrity = runAdminRepairDataIntegrity_(actor, { skipBackup: true });
  const structure = runAdminOptimizeStructure_(actor, { skipBackup: true });
  // Data Integrity already rebuilds the clean analytics ledger. Avoid a second
  // full history scan here; reuse that result and only perform the final audit.
  const analytics = dataIntegrity && dataIntegrity.analyticsLedger ? dataIntegrity.analyticsLedger : null;
  const finalInspection = inspectDatabaseKalibrasiKm_();
  return {
    success: Boolean(finalInspection.ready),
    backup: backup,
    schema: schema,
    workflow: workflow,
    dataIntegrity: dataIntegrity,
    structure: structure,
    analyticsLedger: analytics,
    finalInspection: finalInspection
  };
}

function runAdminDatabaseMaintenance(action, sessionToken, deferHealth) {
  const actor = requireAdminCenterPermission_('ADMIN_SETUP', sessionToken);
  const key = normalizeLookupText_(action).replace(/\s+/g, '_');
  let result;
  if (key === 'AUDIT') {
    result = getAdminDatabaseHealthAudit(sessionToken, true);
  } else if (key === 'REPAIR_TRIGGER_ACCESS') {
    result = runAdminRepairTriggerAccess_(actor);
  } else if (key === 'CLEAN_STRUCTURE') {
    result = cleanDatabaseStructure_(actor, { skipFoundation: true });
  } else {
    throw new Error('Aksi maintenance ini telah dinonaktifkan dari dashboard: ' + key +
      '. Gunakan Audit Sistem, Perbaiki Trigger & Access, atau Bersihkan Row / Column Kosong.');
  }
  const isAudit = key === 'AUDIT';
  if (!isAudit) {
    invalidateAppCaches_();
    invalidateSecurityCaches_();
  }
  logSystemActivity_('', '', 'ADMIN_DATABASE_MAINTENANCE', key, '', '',
    'Maintenance database: ' + key, actor);
  return {
    success: true,
    action: key,
    result: result,
    healthDeferred: !isAudit && deferHealth === true,
    health: isAudit ? result : deferHealth === true ? null : getAdminDatabaseHealthAudit(sessionToken, true)
  };
}

/** Explicit admin maintenance. Never runs during login or normal analytics reads. */
function databaseStructureSpecs_() {
  return [
    [KM_APP.SHEETS.LOG, KM_APP.LOG_HEADERS],
    [UNIT_KM_LOG_SHEET, UNIT_KM_LOG_HEADERS],
    [ANALYTICS_CALIBRATION_SHEET, ANALYTICS_CALIBRATION_HEADERS],
    [UNIT_STATUS_LOG_SHEET, UNIT_STATUS_LOG_HEADERS],
    [UNIT_STATUS_3L_LOG_SHEET, UNIT_STATUS_3L_LOG_HEADERS],
    [MASTER_UNIT_MANAGEMENT.MOVEMENT_SHEET, MASTER_UNIT_MANAGEMENT.MOVEMENT_HEADERS],
    [MASTER_UNIT_MANAGEMENT.REQUEST_SHEET, MASTER_UNIT_MANAGEMENT.REQUEST_HEADERS]
  ];
}

function databaseBlankRowGroups_(values, formulas, notes) {
  const groups = [];
  values.forEach((row, offset) => {
    const occupied = row.some(value => value !== '' && value != null) ||
      (formulas[offset] || []).some(Boolean) || (notes[offset] || []).some(Boolean);
    if (occupied) return;
    const rowNumber = offset + 2;
    const last = groups[groups.length - 1];
    if (last && last.start + last.count === rowNumber) last.count++;
    else groups.push({ start: rowNumber, count: 1 });
  });
  return groups;
}

function removeSessionRowCaches_(sessionIds) {
  const keys = uniqueClean_(sessionIds).map(sessionCacheKey_);
  if (!keys.length) return;
  const cache = CacheService.getScriptCache();
  for (let start = 0; start < keys.length; start += 100) {
    cache.removeAll(keys.slice(start, start + 100));
  }
}

/**
 * Merapatkan tabel auth tanpa deleteRows agar formatting/validation kolom tetap
 * utuh. Fungsi ini hanya boleh dipanggil setelah backup dan di dalam ScriptLock.
 */
function normalizeAuthStorageSheetUnsafe_(sheet, headers, keyHeader, clearSessionCache) {
  ensureHeaders_(sheet, headers);
  const map = getHeaderMap_(sheet);
  const keyColumn = map[keyHeader];
  if (!keyColumn) throw new Error('Header key ' + keyHeader + ' tidak ditemukan di ' + sheet.getName() + '.');
  if (sheet.getLastRow() < 2) {
    return { sheet: sheet.getName(), records: 0, removedGapRows: 0, removedBlankRows: 0 };
  }

  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const values = range.getValues();
  const formulas = range.getFormulas();
  const notes = range.getNotes();
  const keyIndex = keyColumn - 1;
  const records = [];
  const gapRows = [];

  values.forEach((row, offset) => {
    if (cleanText_(row[keyIndex])) records.push(row);
    else gapRows.push(offset + 2);
  });
  if (!gapRows.length) {
    return { sheet: sheet.getName(), records: records.length, removedGapRows: 0, removedBlankRows: 0 };
  }

  const hasFormula = formulas.some(row => row.some(Boolean));
  const hasNote = notes.some(row => row.some(Boolean));
  if (hasFormula || hasNote) {
    throw new Error('Perapihan ' + sheet.getName() + ' dihentikan: ditemukan formula/note yang harus ditinjau manual.');
  }

  const width = sheet.getLastColumn();
  const normalized = records.concat(
    new Array(values.length - records.length).fill(null).map(() => new Array(width).fill(''))
  );
  range.setValues(normalized);
  if (clearSessionCache) {
    removeSessionRowCaches_(records.map(row => cleanText_(row[keyIndex])));
  }
  return {
    sheet: sheet.getName(),
    records: records.length,
    removedGapRows: gapRows.length,
    removedBlankRows: gapRows.length,
    firstGapRow: gapRows[0],
    lastGapRow: gapRows[gapRows.length - 1]
  };
}

function normalizeAuthStorageRowsUnsafe_(ss) {
  const sessions = getOrCreateSheet_(ss, KM_APP.SHEETS.SESSIONS);
  const authLog = getOrCreateSheet_(ss, KM_APP.SHEETS.AUTH_LOG);
  return {
    sessions: normalizeAuthStorageSheetUnsafe_(sessions, KM_AUTH.SESSION_HEADERS, 'SESSION ID', true),
    authLog: normalizeAuthStorageSheetUnsafe_(authLog, KM_AUTH.AUTH_LOG_HEADERS, 'EVENT ID', false)
  };
}

function cleanDatabaseStructure_(actor, options) {
  const settings = options || {};
  const lock = acquireScriptLock_(30000, 'perapihan struktur database');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const plans = databaseStructureSpecs_().map(spec => {
      const sheet = ss.getSheetByName(spec[0]);
      if (!sheet) return null;
      // Fail early on ambiguous headers, before taking a backup or deleting rows.
      if (sheet.getLastRow()) getHistoryHeaderMap_(sheet, spec[1]);
      const range = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()) : null;
      return { sheet: sheet, headers: spec[1], groups: range ? databaseBlankRowGroups_(range.getValues(), range.getFormulas(), range.getNotes()) : [] };
    }).filter(Boolean);
    const backup = settings.skipBackup ? null : createAdministrationBackup_(actor, 'BEFORE_STRUCTURE_CLEANUP');
    const authStorage = normalizeAuthStorageRowsUnsafe_(ss);
    const changes = [authStorage.sessions, authStorage.authLog];
    plans.forEach(plan => {
      let deletedRows = 0;
      plan.groups.slice().reverse().forEach(group => {
        // Re-read exact rows immediately before deletion to preserve concurrent manual input.
        const check = plan.sheet.getRange(group.start, 1, group.count, plan.sheet.getLastColumn());
        const occupied = check.getValues().some(row => row.some(value => value !== '' && value != null)) ||
          check.getFormulas().some(row => row.some(Boolean)) || check.getNotes().some(row => row.some(Boolean));
        if (occupied) throw new Error('Data berubah saat perapihan ' + plan.sheet.getName() + '. Jalankan ulang audit.');
        plan.sheet.deleteRows(group.start, group.count); deletedRows += group.count;
      });
      trimSheetGrid_(plan.sheet, plan.headers.length, 100, 101);
      changes.push({ sheet: plan.sheet.getName(), removedBlankRows: deletedRows });
    });
    const headerOrder = settings.skipFoundation ? [] : standardizeEnterpriseHeaderOrder_(ss);
    if (!settings.skipFoundation) formatDatabaseTables_(ss);
    const compacted = compactDatabaseGrid_();
    SpreadsheetApp.flush();
    invalidateAppCaches_();
    const inventory = auditDatabaseSheetInventory_(ss);
    logSystemActivity_('', '', 'DATABASE_STRUCTURE_CLEANED', '', '', '',
      JSON.stringify({ rows: changes, headerOrder: headerOrder, inventory: inventory }), actor, true);
    return { success: true, backup: backup, sheets: changes, authStorage: authStorage, headerOrder: headerOrder,
      compacted: compacted, inventory: inventory,
      note: settings.skipFoundation
        ? 'Gap KM_SESSIONS/KM_AUTH_LOG, baris kosong log/ledger, dan kapasitas grid kosong dibersihkan; data valid tetap dipertahankan.'
        : 'Gap KM_SESSIONS/KM_AUTH_LOG dan baris kosong log/ledger dibersihkan, urutan header canonical diseragamkan, dan data valid tetap dipertahankan.' };
  } finally { lock.releaseLock(); }
}

function databaseColumnFormat_(header) {
  const key = cleanText_(header).toUpperCase().replace(/\s+/g, ' ');
  if (key === 'TIMESTAMP' || /TIMESTAMP|WAKTU/.test(key) || / AT$/.test(key) || / UNTIL$/.test(key)) return 'dd mmmm yyyy, HH.mm.ss';
  if (/(^| )(TANGGAL|TGL|DATE)( |$)/.test(key)) return 'dd mmmm yyyy';
  if (['KM ODOMETER', 'KALIBRASI KE', 'CALIBRATION SEQUENCE', 'PASSWORD ITERATIONS', 'FAILED LOGIN COUNT', 'SESSION VERSION', 'KM REVISION'].includes(key)) return '#,##0';
  return '@';
}

function formatDatabaseTables_(ss) {
  const names = [KM_APP.SHEETS.CONFIG, KM_APP.SHEETS.MASTER_UNIT, KM_APP.SHEETS.USERS,
    KM_APP.SHEETS.ROLES, KM_APP.SHEETS.SESSIONS, KM_APP.SHEETS.AUTH_LOG,
    LEGACY_UNIT_KM_LOG_SHEET].concat(databaseStructureSpecs_().map(spec => spec[0]));
  const formatted = [];
  const skippedFormats = [];
  Array.from(new Set(names)).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || !sheet.getLastColumn()) return;
    if (sheet.getMaxRows() < 2) sheet.insertRowsAfter(1, 1);
    const width = sheet.getLastColumn(), lastRow = Math.max(2, sheet.getLastRow());
    formatted.push(name);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, width).setBackground('#F0F2F5').setFontColor('#1F2937')
      .setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setWrap(true).setHorizontalAlignment('left');
    sheet.setRowHeight(1, 42);
    sheet.getRange(2, 1, lastRow-1, width).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
    const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0];
    headers.forEach((header, index) => {
      if (!cleanText_(header)) return;
      if (!trySetNumberFormat_(sheet.getRange(2, index+1, Math.max(1,sheet.getMaxRows()-1), 1), databaseColumnFormat_(header))) {
        skippedFormats.push({ sheet: name, column: header, reason: 'Typed column memakai format tabel Google Sheets.' });
      }
      const key = cleanText_(header).toUpperCase();
      const size = /JSON|CATATAN|NOTE|REASON|OPERATING POINT/.test(key) ? 260 :
        /EMAIL| BY| OLEH|NAMA/.test(key) ? 210 : /TIMESTAMP| AT$| UNTIL$/.test(key) ? 165 : 145;
      sheet.setColumnWidth(index+1, size);
    });
    if (name === KM_APP.SHEETS.MASTER_UNIT) sheet.setFrozenColumns(Math.min(2, width));
  });
  return { formattedSheets: formatted, count: formatted.length, skippedFormats: skippedFormats };
}

function getAdminBackupFolder_() {
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty(KM_ADMIN_BACKUP_FOLDER_PROPERTY);
  if (stored) {
    try { return DriveApp.getFolderById(stored); } catch (error) {}
  }
  const folders = DriveApp.getFoldersByName(KM_ADMIN_BACKUP_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(KM_ADMIN_BACKUP_FOLDER_NAME);
  props.setProperty(KM_ADMIN_BACKUP_FOLDER_PROPERTY, folder.getId());
  return folder;
}

function createAdministrationBackup_(actor, label) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folder = getAdminBackupFolder_();
  const source = DriveApp.getFileById(ss.getId());
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const name = KM_ADMIN_BACKUP_PREFIX + stamp + (label ? '_' + String(label).replace(/[^A-Za-z0-9_-]+/g, '_') : '');
  const copy = source.makeCopy(name, folder);
  copy.setDescription(JSON.stringify({ sourceSpreadsheetId: ss.getId(), createdBy: actor && actor.email || '', createdAt: new Date().toISOString() }));
  return { id: copy.getId(), name: copy.getName(), createdAt: copy.getDateCreated().toISOString(), size: copy.getSize() };
}

function createAdministrationBackup(sessionToken) {
  const actor = requireAdminCenterPermission_('ADMIN_SETUP', sessionToken);
  const backup = createAdministrationBackup_(actor, 'MANUAL');
  logSystemActivity_('', '', 'ADMIN_BACKUP_CREATED', 'SPREADSHEET', '', backup.id,
    backup.name, actor);
  return { success: true, backup: backup };
}

function listAdministrationBackups(sessionToken) {
  requireAdminCenterPermission_('ADMIN_SETUP', sessionToken);
  const folder = getAdminBackupFolder_();
  const files = folder.getFiles();
  const rows = [];
  while (files.hasNext()) {
    const file = files.next();
    if (file.isTrashed() || file.getName().indexOf(KM_ADMIN_BACKUP_PREFIX) !== 0) continue;
    rows.push({
      id: file.getId(), name: file.getName(), createdAt: file.getDateCreated().toISOString(),
      updatedAt: file.getLastUpdated().toISOString(), size: file.getSize(), url: file.getUrl()
    });
  }
  rows.sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)));
  return { backups: rows.slice(0, 50), folderId: folder.getId(), folderName: folder.getName() };
}

function adminBackupFileBelongsToFolder_(file, folder) {
  const parents = file.getParents();
  while (parents.hasNext()) if (parents.next().getId() === folder.getId()) return true;
  return false;
}

function restoreAdminSheetData_(sourceSheet, targetSheet) {
  const rows = Math.max(1, sourceSheet.getLastRow());
  const cols = Math.max(1, sourceSheet.getLastColumn());
  if (targetSheet.getMaxRows() < rows) targetSheet.insertRowsAfter(targetSheet.getMaxRows(), rows - targetSheet.getMaxRows());
  if (targetSheet.getMaxColumns() < cols) targetSheet.insertColumnsAfter(targetSheet.getMaxColumns(), cols - targetSheet.getMaxColumns());
  targetSheet.clear({ contentsOnly: false });
  const sourceRange = sourceSheet.getRange(1, 1, rows, cols);
  const values = sourceRange.getValues();
  targetSheet.getRange(1, 1, rows, cols).setValues(values);
  const formulas = sourceRange.getFormulas();
  formulas.forEach((row, r) => row.forEach((formula, c) => {
    if (formula) targetSheet.getRange(r + 1, c + 1).setFormula(formula);
  }));
  targetSheet.setFrozenRows(sourceSheet.getFrozenRows());
  targetSheet.setFrozenColumns(sourceSheet.getFrozenColumns());
  if (sourceSheet.isSheetHidden() && !targetSheet.isSheetHidden()) targetSheet.hideSheet();
  if (!sourceSheet.isSheetHidden() && targetSheet.isSheetHidden()) targetSheet.showSheet();
}

function restoreAdministrationBackup(backupFileId, confirmation, sessionToken) {
  const actor = requireAdminCenterPermission_('ADMIN_SETUP', sessionToken);
  if (String(confirmation || '').trim().toUpperCase() !== 'RESTORE') {
    throw new Error('Konfirmasi recovery harus tepat: RESTORE');
  }
  const folder = getAdminBackupFolder_();
  const file = DriveApp.getFileById(cleanText_(backupFileId));
  if (!adminBackupFileBelongsToFolder_(file, folder) || file.getName().indexOf(KM_ADMIN_BACKUP_PREFIX) !== 0) {
    throw new Error('File bukan backup Kalibrasi KM yang terdaftar.');
  }
  const preRestore = createAdministrationBackup_(actor, 'PRE_RESTORE');
  const backupSs = SpreadsheetApp.openById(file.getId());
  const targetSs = SpreadsheetApp.getActiveSpreadsheet();
  const restored = [];
  KM_ADMIN_RESTORE_SHEETS.forEach(name => {
    const sourceSheet = backupSs.getSheetByName(name);
    if (!sourceSheet) return;
    const targetSheet = targetSs.getSheetByName(name) || targetSs.insertSheet(name);
    restoreAdminSheetData_(sourceSheet, targetSheet);
    restored.push(name);
  });
  SpreadsheetApp.flush();
  setupAuthDatabase_(targetSs);
  setupMasterUnitKmDatabase_();
  removeAutomaticSheetProtections_();
  ensureKalibrasiKmEditTrigger_();
  invalidateAppCaches_();
  invalidateSecurityCaches_();
  logSystemActivity_('', '', 'ADMIN_BACKUP_RESTORED', 'SPREADSHEET', file.getId(), preRestore.id,
    'Restore: ' + file.getName() + ' | pre-restore backup: ' + preRestore.name, actor);
  return {
    success: true,
    restoredSheets: restored,
    sourceBackup: { id: file.getId(), name: file.getName() },
    preRestoreBackup: preRestore,
    note: 'Recovery database selesai. Repository foto odometer di Drive tidak diduplikasi/diubah.'
  };
}


/** Legacy sheet consolidation and schema migration. */
let KM_REQUEST_MIGRATION_DIRECTORY_ = null;
let KM_REQUEST_STATUS3L_OPTIONS_ = null;

function historyMergeFingerprint_(row, map) {
  const historyId = map['HISTORY ID'] ? cleanText_(row[map['HISTORY ID'] - 1]) : '';
  if (historyId) return 'HID|' + historyId;
  const updateId = map['UPDATE ID'] ? cleanText_(row[map['UPDATE ID'] - 1]) : '';
  const nopol = map['NOMOR POLISI'] ? normalizeNopol_(row[map['NOMOR POLISI'] - 1]) : '';
  const sequence = map['KALIBRASI KE'] ? parseCalibrationNumber_(row[map['KALIBRASI KE'] - 1]) : 0;
  const type = map['RECORD TYPE'] ? normalizeLookupText_(row[map['RECORD TYPE'] - 1]) : '';
  const timestamp = map['TIMESTAMP'] ? dateToIso_(row[map['TIMESTAMP'] - 1]) || cleanText_(row[map['TIMESTAMP'] - 1]) : '';
  const photoId = historyOdometerPhotoId_(row, map);
  return ['ROW', updateId, nopol, sequence, type, timestamp, photoId].join('|');
}

function mergeLegacyKmHistoryIntoCanonical_(ss, report) {
  const legacy = ss.getSheetByName(LEGACY_UNIT_KM_LOG_SHEET);
  const canonical = ss.getSheetByName(UNIT_KM_LOG_SHEET);
  if (!legacy || !canonical || legacy.getSheetId() === canonical.getSheetId()) {
    return { merged: 0, skipped: 0, deletedLegacy: false };
  }
  const legacyMap = ensureHistoryHeaders_(legacy, UNIT_KM_LOG_HEADERS);
  const canonicalMap = ensureHistoryHeaders_(canonical, UNIT_KM_LOG_HEADERS);
  const fingerprints = new Set();
  if (canonical.getLastRow() > 1) {
    canonical.getRange(2, 1, canonical.getLastRow() - 1, canonical.getLastColumn()).getValues()
      .forEach(row => fingerprints.add(historyMergeFingerprint_(row, canonicalMap)));
  }
  const additions = [];
  let skipped = 0;
  if (legacy.getLastRow() > 1) {
    legacy.getRange(2, 1, legacy.getLastRow() - 1, legacy.getLastColumn()).getValues().forEach(row => {
      if (!row.some(value => value !== '' && value != null)) return;
      const fingerprint = historyMergeFingerprint_(row, legacyMap);
      if (fingerprints.has(fingerprint)) { skipped++; return; }
      fingerprints.add(fingerprint);
      const entry = {};
      UNIT_KM_LOG_HEADERS.forEach(header => {
        entry[header] = legacyMap[header] ? row[legacyMap[header] - 1] : '';
      });
      additions.push(buildHistoryRow_(canonical, canonicalMap, entry));
    });
  }
  if (additions.length) {
    ensureSheetCapacity_(canonical, additions.length);
    canonical.getRange(canonical.getLastRow() + 1, 1, additions.length, canonical.getLastColumn()).setValues(additions);
  }
  // Backup dibuat oleh caller sebelum migrasi schema terbaru. Setelah merge selesai,
  // legacy dapat dihapus dari workbook aktif agar tidak lagi menjadi dummy source.
  ss.deleteSheet(legacy);
  if (report) {
    report.droppedSheets.push(LEGACY_UNIT_KM_LOG_SHEET);
    report.mergedLegacyHistory = { merged: additions.length, skipped: skipped, target: UNIT_KM_LOG_SHEET };
  }
  return { merged: additions.length, skipped: skipped, deletedLegacy: true };
}

function databaseNeedsMigrationBackup_(ss) {
  if (ss.getSheetByName(LEGACY_UNIT_KM_LOG_SHEET) || ss.getSheetByName('KM_PENGAJUAN')) return true;
  const master = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  if (master && master.getLastColumn()) {
    // Setiap destructive schema cleanup wajib membuat backup lebih dulu. Ini mencakup
    // field legacy lama sekaligus field yang resmi dipensiunkan pada schema v19.
    const obsolete = new Set([
      'KM BITRANS', 'KM BITRACK', 'KM CALIBRATION TICKET', 'TIKET KALIBRASI TERAKHIR',
      'KM CALIBRATION AT', 'WAKTU KALIBRASI TERAKHIR',
      'NOMOR FLEET', 'NO FLEET', 'NO. FLEET', 'FLEET NUMBER', 'NOMOR UNIT', 'NO UNIT', 'KODE UNIT',
      'BRAND', 'MERK UNIT', 'MEREK UNIT', 'MERK', 'MEREK',
      'TAHUN UNIT', 'TAHUN', 'NAMA DRIVER', 'DRIVER', 'PENGEMUDI',
      'NAMA CUSTOMER', 'CUSTOMER', 'CUSTOMER NAME'
    ].map(normalizeHeaderKey_));
    const headers = master.getRange(1,1,1,master.getLastColumn()).getDisplayValues()[0];
    if (headers.some(header => obsolete.has(normalizeHeaderKey_(header)))) return true;
  }
  const config = ss.getSheetByName(KM_APP.SHEETS.CONFIG);
  if (config && config.getLastColumn()) {
    const configHeaders = config.getRange(1,1,1,config.getLastColumn()).getDisplayValues()[0];
    if (configHeaders.some(header => normalizeHeaderKey_(header) === normalizeHeaderKey_('NAMA CUSTOMER'))) return true;
  }
  return false;
}

function consolidateLegacyDatabaseSheets_(actor, options) {
  const settings = options || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backup = settings.skipBackup ? null : createAdministrationBackup_(actor, 'BEFORE_LEGACY_CONSOLIDATION');
  const report = { merged: [], deleted: [], archived: [], backup: backup };
  const canonical = ensureUnitKmHistorySheet_(ss);
  const legacy = ss.getSheetByName(LEGACY_UNIT_KM_LOG_SHEET);
  if (legacy && canonical && legacy.getSheetId() !== canonical.getSheetId()) {
    const mergeReport = { droppedSheets: [], mergedLegacyHistory: null };
    const result = mergeLegacyKmHistoryIntoCanonical_(ss, mergeReport);
    report.merged.push({ source: LEGACY_UNIT_KM_LOG_SHEET, target: UNIT_KM_LOG_SHEET, rows: result.merged, skipped: result.skipped });
    if (result.deletedLegacy) report.deleted.push(LEGACY_UNIT_KM_LOG_SHEET);
  }
  const oldRequest = ss.getSheetByName('KM_PENGAJUAN');
  if (oldRequest) {
    if (oldRequest.getLastRow() <= 1 || oldRequest.getLastColumn() === 0) {
      ss.deleteSheet(oldRequest);
      report.deleted.push('KM_PENGAJUAN');
    } else {
      const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
      const archiveName = ('ARCHIVE_KM_PENGAJUAN_' + stamp).slice(0, 99);
      oldRequest.setName(archiveName);
      if (!oldRequest.isSheetHidden()) oldRequest.hideSheet();
      report.archived.push({ source: 'KM_PENGAJUAN', archive: archiveName, rows: Math.max(0, oldRequest.getLastRow() - 1) });
    }
  }
  SpreadsheetApp.flush();
  report.inventory = auditDatabaseSheetInventory_(ss);
  logSystemActivity_('', '', 'DATABASE_LEGACY_CONSOLIDATED', '', '', '', JSON.stringify(report), actor, true);
  return report;
}

function migrateMonitoringDatabase_(ss) {
  const report = {
    schemaVersion: KM_DATABASE_SCHEMA_VERSION,
    droppedSheets: [],
    renamedSheets: [],
    removedHeaders: {},
    createdSheets: [],
    configMigrated: false
  };

  // 1. Snapshot MASTER_CONFIG sebelum layout legacy dibersihkan.
  const config = ss.getSheetByName(KM_APP.SHEETS.CONFIG);
  KM_REQUEST_MIGRATION_DIRECTORY_ = [];
  KM_REQUEST_STATUS3L_OPTIONS_ = [];
  if (config && config.getLastRow() >= 1 && config.getLastColumn() >= 1) {
    const headers = config.getRange(1, 1, 1, config.getLastColumn()).getDisplayValues()[0];
    const opIndex = headers.findIndex(h => normalizeHeaderKey_(h) === normalizeHeaderKey_('OPERATING POINT'));
    const opcIndex = headers.findIndex(h => normalizeHeaderKey_(h) === normalizeHeaderKey_('NAMA OPC/WHC'));
    const tohIndex = headers.findIndex(h => normalizeHeaderKey_(h) === normalizeHeaderKey_('NAMA TOH/WOH'));
    const status3LIndex = headers.findIndex(h => normalizeHeaderKey_(h) === normalizeHeaderKey_('STATUS PROSES 3L'));
    const values = config.getLastRow() > 1
      ? config.getRange(2, 1, config.getLastRow() - 1, config.getLastColumn()).getDisplayValues()
      : [];

    const directoryRows = Math.max(
      opIndex >= 0 ? values.reduce((m, row, i) => cleanText_(row[opIndex]) ? Math.max(m, i + 1) : m, 0) : 0,
      opcIndex >= 0 ? values.reduce((m, row, i) => cleanText_(row[opcIndex]) ? Math.max(m, i + 1) : m, 0) : 0,
      tohIndex >= 0 ? values.reduce((m, row, i) => cleanText_(row[tohIndex]) ? Math.max(m, i + 1) : m, 0) : 0
    );
    for (let i = 0; i < directoryRows; i++) {
      KM_REQUEST_MIGRATION_DIRECTORY_.push([
        opIndex >= 0 ? safeSheetText_(values[i] && values[i][opIndex]) : '',
        opcIndex >= 0 ? safeSheetText_(values[i] && values[i][opcIndex]) : '',
        tohIndex >= 0 ? safeSheetText_(values[i] && values[i][tohIndex]) : ''
      ]);
    }
    if (status3LIndex >= 0) {
      KM_REQUEST_STATUS3L_OPTIONS_ = uniqueClean_(values.map(row => row[status3LIndex]));
    }
    report.configMigrated = true;
  }

  // 2. MASTER_UNIT: hapus field lama yang tidak lagi dipakai flow monitoring terbaru.
  const masterObsolete = new Set([
    'KM BITRANS', 'KM BITRACK', 'KM CALIBRATION TICKET',
    'TIKET KALIBRASI TERAKHIR', 'KM CALIBRATION AT', 'WAKTU KALIBRASI TERAKHIR',
    // Schema v19 — field tidak digunakan dan seluruh aliasnya dibersihkan dari MASTER_UNIT.
    'NOMOR FLEET', 'NO FLEET', 'NO. FLEET', 'FLEET NUMBER', 'NOMOR UNIT', 'NO UNIT', 'KODE UNIT',
    'BRAND', 'MERK UNIT', 'MEREK UNIT', 'MERK', 'MEREK',
    'TAHUN UNIT', 'TAHUN', 'NAMA DRIVER', 'DRIVER', 'PENGEMUDI',
    'NAMA CUSTOMER', 'CUSTOMER', 'CUSTOMER NAME'
  ].map(normalizeHeaderKey_));
  const master = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  if (master && master.getLastColumn()) {
    const removed = [];
    const headers = master.getRange(1, 1, 1, master.getLastColumn()).getDisplayValues()[0];
    for (let i = headers.length - 1; i >= 0; i--) {
      if (masterObsolete.has(normalizeHeaderKey_(headers[i]))) {
        removed.push(headers[i]);
        master.deleteColumn(i + 1);
      }
    }
    if (removed.length) report.removedHeaders[KM_APP.SHEETS.MASTER_UNIT] = removed.reverse();
  }

  // 3. HISTORY_KALIBRASI_KM: migrasikan ledger lama dan bersihkan field pembanding usang.
  let history = ensureUnitKmHistorySheet_(ss, report);
  // Bila canonical dan legacy sama-sama ada, merge row unik ke canonical lalu
  // hapus legacy dari workbook aktif. Backup dilakukan oleh setup core sebelum migrasi.
  if (ss.getSheetByName(LEGACY_UNIT_KM_LOG_SHEET) && history && history.getName() === UNIT_KM_LOG_SHEET) {
    mergeLegacyKmHistoryIntoCanonical_(ss, report);
    history = ss.getSheetByName(UNIT_KM_LOG_SHEET);
  }
  if (history && history.getLastColumn()) {
    const removed = [];
    const obsoleteHistory = new Set(['KM BITRANS', 'KM BITRACK'].map(normalizeHeaderKey_));
    const headers = history.getRange(1, 1, 1, history.getLastColumn()).getDisplayValues()[0];
    for (let i = headers.length - 1; i >= 0; i--) {
      if (obsoleteHistory.has(normalizeHeaderKey_(headers[i]))) {
        removed.push(headers[i]);
        history.deleteColumn(i + 1);
      }
    }
    if (removed.length) report.removedHeaders[UNIT_KM_LOG_SHEET] = removed.reverse();

    const map = ensureHistoryHeaders_(history, UNIT_KM_LOG_HEADERS);
    if (history.getLastRow() > 1) {
      const rows = history.getRange(2, 1, history.getLastRow() - 1, history.getLastColumn()).getValues();
      history.getRange(2, map['FOTO DIUNGGAH OLEH'], rows.length, 1).setValues(rows.map(row => [
        safeSheetText_(row[map['FOTO DIUNGGAH OLEH'] - 1] ||
          (row[map['FOTO ODOMETER ID'] - 1] ? row[map['USER EMAIL'] - 1] : ''))
      ]));
      if (map['DATA SEBELUM']) {
        history.getRange(2, map['DATA SEBELUM'], rows.length, 1).setValues(rows.map(row => {
          try {
            const old = JSON.parse(row[map['DATA SEBELUM'] - 1] || '{}');
            const clean = {};
            ['statusUnit', 'status3L', 'status3LNote', 'status3LUpdatedAt',
              'status3LUpdatedBy', 'calibrationNumber', 'kmOdometer', 'kmDate',
              'kmUpdatedAt', 'kmUpdatedBy', 'kmPhotoId', 'kmPhotoAddedAt',
              'kmPhotoAddedBy'].forEach(key => {
              if (old[key] !== undefined) clean[key] = old[key];
            });
            return [Object.keys(clean).length ? JSON.stringify(clean) : ''];
          } catch (error) {
            return [''];
          }
        }));
      }
    }
  }

  // 4. KM_ACTIVITY_LOG: hapus referensi NO TIKET karena flow terbaru berbasis unit.
  const activityLog = ss.getSheetByName(KM_APP.SHEETS.LOG);
  if (activityLog && activityLog.getLastColumn()) {
    const headers = activityLog.getRange(1, 1, 1, activityLog.getLastColumn()).getDisplayValues()[0];
    const removed = [];
    for (let i = headers.length - 1; i >= 0; i--) {
      if (normalizeHeaderKey_(headers[i]) === normalizeHeaderKey_('NO TIKET')) {
        removed.push(headers[i]);
        activityLog.deleteColumn(i + 1);
      }
    }
    if (removed.length) report.removedHeaders[KM_APP.SHEETS.LOG] = removed.reverse();
  }

  // 5. Workflow pengajuan lama sudah tidak digunakan.
  const legacy = ss.getSheetByName('KM_PENGAJUAN');
  if (legacy && ss.getSheets().length > 1) {
    if (legacy.getLastRow() <= 1 || legacy.getLastColumn() === 0) {
      ss.deleteSheet(legacy);
      report.droppedSheets.push('KM_PENGAJUAN');
    } else {
      const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
      const archiveName = ('ARCHIVE_KM_PENGAJUAN_' + stamp).slice(0, 99);
      legacy.setName(archiveName);
      if (!legacy.isSheetHidden()) legacy.hideSheet();
      if (!Array.isArray(report.archivedSheets)) report.archivedSheets = [];
      report.archivedSheets.push({ source: 'KM_PENGAJUAN', archive: archiveName, rows: Math.max(0, legacy.getLastRow() - 1) });
    }
  }

  // 6. Pastikan history Status 3L tersedia.
  if (!ss.getSheetByName(UNIT_STATUS_3L_LOG_SHEET)) {
    const status3LHistory = getOrCreateSheet_(ss, UNIT_STATUS_3L_LOG_SHEET);
    ensureHistoryHeaders_(status3LHistory, UNIT_STATUS_3L_LOG_HEADERS);
    report.createdSheets.push(UNIT_STATUS_3L_LOG_SHEET);
  }

  return report;
}


/** Generic append-only history helpers and unit-coverage endpoints. */
function ensureHistoryHeaders_(sheet, expected) {
  if (!sheet.getLastRow()) {
    if (expected.length > sheet.getMaxColumns()) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), expected.length - sheet.getMaxColumns());
    }
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
    return getHistoryHeaderMap_(sheet, expected);
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(cleanText_);
  const normalized = headers.map(normalizeHeaderKey_);
  const duplicate = expected.find(header => normalized.filter(key => key === normalizeHeaderKey_(header)).length > 1);
  if (duplicate) {
    throw new Error('Header ganda ' + duplicate + ' pada ' + sheet.getName() + '. Hapus salah satu header duplikat sebelum memperbarui unit.');
  }

  const missing = expected.filter(header => !normalized.includes(normalizeHeaderKey_(header)));
  if (missing.length) {
    const firstColumn = sheet.getLastColumn() + 1;
    const requiredColumns = sheet.getLastColumn() + missing.length;
    if (requiredColumns > sheet.getMaxColumns()) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
    }
    sheet.getRange(1, firstColumn, 1, missing.length).setValues([missing]);
  }
  return getHistoryHeaderMap_(sheet, expected);
}

function getHistoryHeaderMap_(sheet, expected) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const normalized = headers.map(normalizeHeaderKey_);
  const map = {};
  expected.forEach(header => {
    const key = normalizeHeaderKey_(header);
    const matches = normalized.reduce((result, candidate, index) => {
      if (candidate === key) result.push(index + 1);
      return result;
    }, []);
    if (matches.length !== 1) {
      throw new Error('Struktur ' + sheet.getName() + ' tidak valid pada header ' + header + '.');
    }
    map[header] = matches[0];
  });
  return map;
}

function historyHeadersReady_(sheet, expected) {
  if (!sheet || !sheet.getLastRow()) return false;
  try {
    getHistoryHeaderMap_(sheet, expected);
    return true;
  } catch (error) {
    return false;
  }
}

function buildHistoryRow_(sheet, map, values) {
  const row = new Array(sheet.getLastColumn()).fill('');
  Object.keys(values).forEach(header => {
    if (map[header]) row[map[header] - 1] = values[header];
  });
  return row;
}

function readRecentHistoryRows_(sheet, nopolColumn, nopol, limit) {
  const target = normalizeNopol_(nopol);
  const matches = [];
  const batchSize = 500;
  const lastRow = sheet.getLastRow();
  if (!target || lastRow < 2 || !(limit > 0)) return matches;

  const plateRange = sheet.getRange(2, nopolColumn, lastRow - 1, 1);
  let rowNumbers = [];
  try {
    const pattern = normalizedNopolTextFinderPattern_(target);
    rowNumbers = plateRange.createTextFinder(pattern)
      .useRegularExpression(true)
      .matchCase(false)
      .matchEntireCell(true)
      .findAll()
      .map(range => range.getRow())
      .sort((a, b) => b - a)
      .slice(0, limit);
  } catch (error) {
    // Fallback kompatibilitas bila TextFinder/regex tidak tersedia pada runtime.
    console.warn('TextFinder history fallback ke scan kolom:', error && error.message ? error.message : error);
    const plates = plateRange.getDisplayValues();
    for (let index = plates.length - 1; index >= 0 && rowNumbers.length < limit; index--) {
      if (normalizeNopol_(plates[index][0]) === target) rowNumbers.push(index + 2);
    }
  }

  if (!rowNumbers.length) return matches;
  const blocks = new Map();
  rowNumbers.forEach(rowNumber => {
    const block = Math.floor((rowNumber - 2) / batchSize);
    if (!blocks.has(block)) blocks.set(block, []);
    blocks.get(block).push(rowNumber);
  });
  blocks.forEach((numbers, block) => {
    const startRow = block * batchSize + 2;
    const endRow = Math.min(lastRow, startRow + batchSize - 1);
    const rows = sheet.getRange(startRow, 1, endRow - startRow + 1, sheet.getLastColumn()).getValues();
    numbers.forEach(rowNumber => {
      const row = rows[rowNumber - startRow];
      // A direct spreadsheet edit can happen between TextFinder and block read.
      if (row && normalizeNopol_(row[nopolColumn - 1]) === target) matches.push(row);
    });
  });
  return matches;
}

function getManagedOperatingPoints_() {
  requirePermission_('MANAGE_USERS');
  // Gunakan directory berbasis header yang sama dengan Area & Coverage, Add/Edit,
  // Transfer OP, dan Data Quality. Ini menghilangkan dependency lama ke kolom A.
  return uniqueClean_(effectiveMasterUnitCoverageDirectory_().map(item => item.operatingPoint));
}
/** Revision covers the editable directory fields in their displayed row order. */
function coverageDirectoryRevision_(rows) {
  const content = JSON.stringify((rows || []).map(item => [item.operatingPoint || '', item.opc || '', item.toh || '']));
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, content));
}

function unitCoverageRevision_(unit) {
  return JSON.stringify([unit.unitId, unit.operatingPoint, unit.toh, unit.masterStatus, unit.lastMovementId]);
}

/** Page-sized response: never sends out-of-scope units or user-directory details. */
function getAdminUnitCoverageData(filters, sessionToken) {
  const actor = requireAnyMasterUnitManagementPermission_(sessionToken);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  if (!sheet || !masterUnitManagementSchemaReady_(sheet)) throw new Error('Jalankan Repair Struktur Master Unit terlebih dahulu.');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
  const values = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues() : [];
  masterUnitCoverageDirectory_({ indexes: indexes, rows: values });
  const input = filters || {};
  const query = cleanText_(input.query).toLowerCase().slice(0, 120);
  const opKey = coverageOperatingPointKey_(input.operatingPoint);
  const groups = new Map();
  let units = [];
  values.forEach(row => {
    if (!cleanText_(row[indexes.nopol])) return;
    const operatingPoint = cleanText_(row[indexes.operatingPoint]);
    if (!canAccessOperatingPoint_(actor, operatingPoint)) return;
    const snapshot = masterUnitRecordSnapshot_(row, indexes);
    const key = coverageOperatingPointKey_(operatingPoint);
    if (!groups.has(key)) groups.set(key, { operatingPoint: operatingPoint, total: 0, active: 0, unmapped: 0 });
    const group = groups.get(key);
    group.total++;
    if (normalizeLookupText_(snapshot.masterStatus) === 'ACTIVE') group.active++;
    if (!operatingPoint || !snapshot.toh) group.unmapped++;
    units.push({ unitId: snapshot.unitId, nopol: snapshot.nopol, unitType: snapshot.unitType,
      operatingPoint: operatingPoint, toh: snapshot.toh,
      masterStatus: snapshot.masterStatus, revision: unitCoverageRevision_(snapshot) });
  });
  const totalUnits = units.length;
  units = units.filter(unit => (!opKey || coverageOperatingPointKey_(unit.operatingPoint) === opKey) &&
    (!input.unmapped || !unit.operatingPoint || !unit.toh) &&
    (!query || [unit.nopol, unit.unitId, unit.unitType, unit.operatingPoint, unit.toh].join(' ').toLowerCase().includes(query)));
  units.sort((a,b) => a.nopol.localeCompare(b.nopol, 'id', { numeric: true }));
  const pageSize = 50;
  const pages = Math.max(1, Math.ceil(units.length / pageSize));
  const page = Math.min(pages, Math.max(1, Math.floor(Number(input.page) || 1)));
  return { rows: units.slice((page - 1) * pageSize, page * pageSize), total: units.length, totalUnits: totalUnits,
    page: page, pages: pages, pageSize: pageSize,
    groups: Array.from(groups.values()).sort((a,b) => a.operatingPoint.localeCompare(b.operatingPoint)),
    coverage: effectiveMasterUnitCoverageDirectory_().filter(item => canAccessOperatingPoint_(actor, item.operatingPoint)),
    canManage: actor.permissions.includes('MANAGE_MASTER_UNIT'),
    canRequest: actor.permissions.includes('REQUEST_MASTER_UNIT_CHANGE') && masterUnitApprovalEnabled_(),
    globalScope: Boolean(actor.globalScope), loadedAt: new Date().toISOString() };
}

/** Source and destination scope are checked again under lock, with optimistic concurrency. */
function saveAdminUnitCoverageMapping(payload, sessionToken) {
  const actor = requireAnyMasterUnitManagementPermission_(sessionToken);
  const manage = actor.permissions.includes('MANAGE_MASTER_UNIT');
  if (!manage && !actor.permissions.includes('REQUEST_MASTER_UNIT_CHANGE')) throw new Error('ACCESS_DENIED: Tidak memiliki izin pemetaan unit.');
  const data = payload || {};
  if (!cleanText_(data.unitId) || !cleanText_(data.expectedRevision)) throw new Error('Pilih unit dari data terbaru terlebih dahulu.');
  const lock = acquireScriptLock_(20000, 'Pemetaan coverage unit');
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
    const before = findMasterUnitManagementTarget_(sheet, data.unitId).snapshot;
    const proposed = validateMasterUnitBasePayload_(data, 'TRANSFER');
    if (!canAccessOperatingPoint_(actor, before.operatingPoint) || !canAccessOperatingPoint_(actor, proposed.operatingPoint)) {
      throw new Error('ACCESS_DENIED: Operating Point asal atau tujuan berada di luar scope akun Anda.');
    }
    if (unitCoverageRevision_(before) !== data.expectedRevision) throw new Error('DATA_CONFLICT: Mapping unit sudah berubah. Muat ulang dan pilih unit kembali.');
    if (normalizeLookupText_(before.masterStatus) !== 'ACTIVE') throw new Error('Unit INACTIVE harus diaktifkan sebelum dipetakan.');
    if (coverageOperatingPointKey_(before.operatingPoint) === coverageOperatingPointKey_(proposed.operatingPoint) &&
      coverageTohKey_(before.toh) === coverageTohKey_(proposed.toh)) throw new Error('Operating Point dan TOH tujuan sama dengan mapping saat ini.');
    if (!manage) {
      const pending = readMasterUnitChangeRequests_(1000).some(item => item.unitId === before.unitId && normalizeLookupText_(item.status) === 'PENDING');
      if (pending) throw new Error('Unit masih memiliki pengajuan PENDING. Selesaikan pengajuan tersebut terlebih dahulu.');
      const result = submitAdminMasterUnitChangeRequest(Object.assign({}, data, { action: 'TRANSFER' }), sessionToken);
      return { success: true, requested: true, requestId: result.requestId };
    }
    const result = applyMasterUnitManagementAction_('TRANSFER', data, actor, { source: 'COVERAGE_MAPPING' });
    logSystemActivity_(before.nopol, '', 'MASTER_UNIT_COVERAGE_MAPPED', 'OPERATING POINT / TOH',
      before.operatingPoint + ' / ' + before.toh, proposed.operatingPoint + ' / ' + proposed.toh, proposed.reason, actor, true);
    return { success: true, requested: false, result: result };
  } finally {
    invalidateAppCaches_();
    lock.releaseLock();
  }
}

function getAdminMasterUnitQuality(sessionToken) {
  const actor = requireAnyMasterUnitManagementPermission_(sessionToken);
  if (!actor.permissions.includes('MANAGE_MASTER_UNIT') && !actor.permissions.includes('APPROVE_MASTER_UNIT_CHANGE')) throw new Error('ACCESS_DENIED: Tidak memiliki izin audit master.');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  if (!sheet) throw new Error('MASTER_UNIT belum tersedia.');
  const indexes = buildAliasIndex_(sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0], MASTER_UNIT_FIELD_ALIASES);
  return masterUnitDataQuality_(sheet, indexes);
}