/** Analytics, data-quality reporting, and unit summary endpoints. */
function analyticsDateValue_(value) {
  return parseDateTimeInput_(value) || parseDateInput_(value) || '';
}

function analyticsStartOfDay_(date) {
  const value = analyticsDateValue_(date);
  if (!value) return '';
  return new Date(value.getFullYear(), value.getMonth(), value.getDate(), 0, 0, 0, 0);
}

function analyticsEndOfDay_(date) {
  const value = analyticsDateValue_(date);
  if (!value) return '';
  return new Date(value.getFullYear(), value.getMonth(), value.getDate(), 23, 59, 59, 999);
}

function analyticsResolveRange_(filters) {
  const input = filters || {};
  const today = analyticsStartOfDay_(new Date());
  let end = analyticsEndOfDay_(input.endDate || today);
  if (!end) end = analyticsEndOfDay_(today);
  let start = analyticsStartOfDay_(input.startDate);
  if (!start) {
    start = new Date(end.getFullYear(), end.getMonth(), end.getDate() - 29, 0, 0, 0, 0);
  }
  if (start.getTime() > end.getTime()) {
    const tmp = start;
    start = analyticsStartOfDay_(end);
    end = analyticsEndOfDay_(tmp);
  }
  if ((input.startDate && !analyticsStartOfDay_(input.startDate)) ||
      (input.endDate && !analyticsEndOfDay_(input.endDate))) {
    throw new Error('Tanggal analisa tidak valid. Gunakan tanggal kalender yang benar.');
  }
  if (input.startDate && input.endDate && String(input.startDate) > String(input.endDate)) {
    throw new Error('Tanggal mulai harus sebelum atau sama dengan tanggal selesai.');
  }
  if ((end.getTime() - start.getTime()) / 86400000 >= 3660) {
    throw new Error('Rentang maksimal 10 tahun. Pilih rentang yang lebih pendek.');
  }
  return { start: start, end: end, startKey: dateOnlyKey_(start), endKey: dateOnlyKey_(end) };
}

function analyticsInRange_(date, range) {
  const value = analyticsDateValue_(date);
  return Boolean(value && value.getTime() >= range.start.getTime() && value.getTime() <= range.end.getTime());
}

function analyticsBucketKey_(date, grain) {
  const value = analyticsDateValue_(date);
  if (!value) return '';
  const normalized = normalizeLookupText_(grain || 'DAILY');
  if (normalized === 'MONTHLY') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  if (normalized === 'WEEKLY') {
    const local = new Date(value.getFullYear(), value.getMonth(), value.getDate(), 12, 0, 0, 0);
    const daysSinceMonday = (local.getDay() + 6) % 7;
    local.setDate(local.getDate() - daysSinceMonday);
    return Utilities.formatDate(local, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function analyticsBuildBucketKeys_(range, grain) {
  const result = [];
  const normalized = normalizeLookupText_(grain || 'DAILY');
  const cursor = new Date(range.start.getFullYear(), range.start.getMonth(), range.start.getDate(), 12, 0, 0, 0);
  if (normalized === 'WEEKLY') {
    cursor.setDate(cursor.getDate() - ((cursor.getDay() + 6) % 7));
  } else if (normalized === 'MONTHLY') {
    cursor.setDate(1);
  }
  const end = range.end.getTime();
  const seen = Object.create(null);
  while (cursor.getTime() <= end) {
    const key = analyticsBucketKey_(cursor, normalized);
    if (key && !seen[key]) { seen[key] = true; result.push(key); }
    if (normalized === 'MONTHLY') cursor.setMonth(cursor.getMonth() + 1, 1);
    else if (normalized === 'WEEKLY') cursor.setDate(cursor.getDate() + 7);
    else cursor.setDate(cursor.getDate() + 1);
  }
  return result;
}

function analyticsLeadHours_(openedAt, completedAt) {
  const opened = analyticsDateValue_(openedAt);
  const completed = analyticsDateValue_(completedAt);
  if (!opened || !completed || completed.getTime() < opened.getTime()) return null;
  return (completed.getTime() - opened.getTime()) / 3600000;
}

function analyticsRound_(value, digits) {
  const numeric = Number(value);
  if (!isFinite(numeric)) return 0;
  const factor = Math.pow(10, Number(digits) || 0);
  return Math.round(numeric * factor) / factor;
}

function analyticsPercent_(numerator, denominator) {
  const total = Number(denominator) || 0;
  return total ? analyticsRound_((Number(numerator) || 0) * 100 / total, 1) : 0;
}

function analyticsNormalizeNote_(value) {
  const text = cleanText_(value).replace(/\s+/g, ' ');
  if (!text) return 'Tanpa catatan';
  return text.length > 80 ? text.slice(0, 77) + '…' : text;
}

function analyticsUnitKey_(nopol, calibrationNumber) {
  const plate = normalizeNopol_(nopol);
  const sequence = parseCalibrationNumber_(calibrationNumber);
  return plate && sequence ? plate + '|' + sequence : '';
}

// Analytics accepts optional legacy columns, while write endpoints keep strict schema validation.
function analyticsReadHistory_(sheet, expected, diagnostics) {
  if (!sheet || sheet.getLastRow() < 2) return { map: {}, rawMap: {}, rows: [] };
  const width = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0];
  const rawMap = {};
  headers.forEach((header, index) => {
    const key = String(header || '').trim();
    if (key) rawMap[key] = index + 1;
  });
  const normalized = headers.map(normalizeHeaderKey_);
  const map = {};
  expected.forEach(header => {
    const matches = normalized.reduce((list, key, i) => {
      if (key === normalizeHeaderKey_(header)) list.push(i + 1);
      return list;
    }, []);
    if (matches.length > 1) throw new Error('Header ganda ' + header + ' di ' + sheet.getName() + '. Perbaiki database sebelum analisa.');
    if (matches.length) map[header] = matches[0];
  });
  if (!map['KALIBRASI KE'] && map['CALIBRATION SEQUENCE']) map['KALIBRASI KE'] = map['CALIBRATION SEQUENCE'];
  ['NOMOR POLISI', 'TIMESTAMP'].forEach(header => {
    if (!map[header]) throw new Error('Header ' + header + ' tidak ditemukan di ' + sheet.getName() + '.');
  });
  const missing = expected.filter(header => !map[header]);
  if (missing.length) diagnostics.warnings.push(sheet.getName() + ': ' + missing.length + ' kolom opsional belum tersedia; histori tetap dibaca.');
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues();
  diagnostics.sources.push({ name: sheet.getName(), rows: rows.length, missingColumns: missing });
  return { map: map, rawMap: rawMap, rows: rows };
}


function analyticsReadCleanLedger_(sheet, diagnostics) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const source = analyticsReadHistory_(sheet, ANALYTICS_CALIBRATION_HEADERS, diagnostics);
  const map = source.map;
  const at = header => map[header] ? map[header] - 1 : -1;
  const seen = Object.create(null);
  const result = [];
  source.rows.forEach((row, offset) => {
    const nopol = at('NOMOR POLISI') >= 0 ? normalizeDisplayNopol_(row[at('NOMOR POLISI')]) : '';
    const sequence = at('CALIBRATION SEQUENCE') >= 0
      ? parseCalibrationNumber_(row[at('CALIBRATION SEQUENCE')])
      : (at('KALIBRASI KE') >= 0 ? parseCalibrationNumber_(row[at('KALIBRASI KE')]) : 0);
    const key = analyticsUnitKey_(nopol, sequence);
    if (!key) return;
    if (seen[key]) throw new Error('ANALYTICS_KALIBRASI_KM memiliki duplicate NOPOL + CALIBRATION SEQUENCE: ' + nopol + ' / ' + sequence + '.');
    seen[key] = true;
    const openedAt = at('TIMESTAMP') >= 0 ? analyticsDateValue_(row[at('TIMESTAMP')]) : '';
    const status3L = at('STATUS 3L') >= 0 ? displayValue_(row[at('STATUS 3L')]) : '';
    const completedAt = at('COMPLETED AT') >= 0 ? analyticsDateValue_(row[at('COMPLETED AT')]) : '';
    result.push({
      key: key,
      rowNumber: offset + 2,
      nopol: nopol,
      calibrationNumber: sequence,
      openedAt: openedAt,
      openedBy: at('USER EMAIL') >= 0 ? displayValue_(row[at('USER EMAIL')]) : '',
      operatingPoint: at('OPERATING POINT') >= 0 ? displayValue_(row[at('OPERATING POINT')]) : '',
      toh: at('NAMA TOH/WOH') >= 0 ? displayValue_(row[at('NAMA TOH/WOH')]) : '',
      statusUnit: at('STATUS UNIT') >= 0 ? displayValue_(row[at('STATUS UNIT')]) : '',
      status3L: status3L,
      status3LInitial: at('STATUS 3L AWAL') >= 0 ? displayValue_(row[at('STATUS 3L AWAL')]) : '',
      status3LUpdatedAt: at('STATUS 3L UPDATED AT') >= 0 ? analyticsDateValue_(row[at('STATUS 3L UPDATED AT')]) : '',
      status3LUpdatedBy: normalizeDoneStatus3LActor_(status3L,
        at('STATUS 3L UPDATED BY') >= 0 ? displayValue_(row[at('STATUS 3L UPDATED BY')]) : ''),
      completedAt: completedAt,
      completedBy: normalizeDoneStatus3LActor_(status3L,
        (at('COMPLETED BY') >= 0 ? displayValue_(row[at('COMPLETED BY')]) : '') ||
        (at('STATUS 3L UPDATED BY') >= 0 ? displayValue_(row[at('STATUS 3L UPDATED BY')]) : '')),
      kmDate: at('TANGGAL KM') >= 0 ? row[at('TANGGAL KM')] : '',
      kmOdometer: at('KM ODOMETER') >= 0 ? parseUnitKm_(row[at('KM ODOMETER')]) : null,
      photoId: at('FOTO ODOMETER ID') >= 0 ? normalizeOdometerPhotoId_(row[at('FOTO ODOMETER ID')]) : '',
      requestId: at('REQUEST ID') >= 0 ? displayValue_(row[at('REQUEST ID')]) : '',
      historyId: at('HISTORY ID') >= 0 ? displayValue_(row[at('HISTORY ID')]) : ''
    });
  });
  return result;
}

function analyticsPreviousRange_(range) {
  const days = Math.round((analyticsStartOfDay_(range.end) - range.start) / 86400000) + 1;
  const end = new Date(range.start.getTime() - 1);
  const start = new Date(range.start);
  start.setDate(start.getDate() - days);
  return { start: start, end: end, startKey: dateOnlyKey_(start), endKey: dateOnlyKey_(end) };
}

function analyticsPeriodMetrics_(submissions, finals, events, range, slaHours) {
  const opened = submissions.filter(item => analyticsInRange_(item.openedAt, range));
  const done = finals.filter(item => analyticsInRange_(item.completedAt, range));
  const leads = done.map(item => item.leadHours).filter(value => value != null && isFinite(value));
  return {
    submissions: opened.length, done: done.length,
    kendala: events.filter(item => analyticsInRange_(item.timestamp, range)).length,
    averageLeadHours: leads.length ? analyticsRound_(leads.reduce((a,b) => a+b, 0) / leads.length, 1) : null,
    slaAchievement: leads.length ? analyticsPercent_(leads.filter(value => value <= slaHours).length, leads.length) : null,
    slaSamples: leads.length
  };
}

// Persentase pengajuan memakai grain unit, bukan jumlah event. Denominatornya
// adalah snapshot unit aktif unik pada scope akun/filter saat analytics dibuka.
// Dengan demikian satu unit yang mengajukan beberapa kali dalam satu periode
// tetap hanya dihitung satu kali pada persentase, sementara volume pengajuan
// tetap tersedia sebagai metrik terpisah.
function analyticsSubmissionPeriodMetrics_(submissions, range, activeUnitByPlate) {
  const opened = submissions.filter(item =>
    analyticsInRange_(item.openedAt, range) && activeUnitByPlate[normalizeNopol_(item.nopol)]
  );
  const submitted = Object.create(null);
  opened.forEach(item => {
    const plateKey = normalizeNopol_(item.nopol);
    if (plateKey && activeUnitByPlate[plateKey]) submitted[plateKey] = true;
  });
  const totalUnits = Object.keys(activeUnitByPlate).length;
  const submittedUnits = Object.keys(submitted).length;
  return {
    submissions: opened.length,
    submittedUnits: submittedUnits,
    unsubmittedUnits: Math.max(0, totalUnits - submittedUnits),
    submittedPercent: analyticsPercent_(submittedUnits, totalUnits)
  };
}

function analyticsComparison_(current, previous, previousRange) {
  const metrics = Object.keys(current).filter(key => key !== 'slaSamples').map(key => {
    const a = current[key], b = previous[key];
    const delta = a == null || b == null ? null : analyticsRound_(a - b, 1);
    return { key: key, current: a, previous: b, delta: delta,
      percentChange: delta == null || b === 0 ? null : analyticsRound_(delta * 100 / Math.abs(b), 1) };
  });
  return { previousStart: previousRange.startKey, previousEnd: previousRange.endKey, metrics: metrics };
}

function analyticsCacheKey_(actor, input, range, grain, slaHours) {
  let revision = getCachedJson_('analytics-revision');
  if (!revision) { revision = Utilities.getUuid(); putCachedJson_('analytics-revision', revision, 600); }
  return 'analytics-integrity-v4:' + sha256Hex_(JSON.stringify([revision, actor.email, actor.roles,
    actor.globalScope, actor.operatingPointKeys || actor.operatingPoints, range.startKey, range.endKey,
    grain, slaHours, normalizeLookupText_(input.operatingPoint), normalizeLookupText_(input.toh)]));
}

function getCalibrationAnalytics(filters, sessionToken) {
  const actor = requirePermission_('VIEW_MONITORING', '', { token: sessionToken });
  const input = filters || {};
  const range = analyticsResolveRange_(input);
  const grain = ['DAILY','WEEKLY','MONTHLY'].includes(normalizeLookupText_(input.grain))
    ? normalizeLookupText_(input.grain) : 'DAILY';
  const slaHours = [2,4,8,12,24,48].includes(Number(input.slaHours)) ? Number(input.slaHours) : 8;
  const opFilter = normalizeLookupText_(input.operatingPoint);
  const tohFilter = normalizeLookupText_(input.toh);
  const cacheKey = analyticsCacheKey_(actor, input, range, grain, slaHours);
  const cached = input.forceRefresh ? null : getCachedJson_(cacheKey);
  if (cached) return cached;
  const diagnostics = { sources: [], warnings: [], invalidDateRows: 0, missingSequenceRows: 0,
    recoveredSubmissions: 0, recoveredCompletions: 0 };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(KM_APP.SHEETS.MASTER_UNIT);
  const cleanLedgerSheet = ss.getSheetByName(ANALYTICS_CALIBRATION_SHEET);
  const cleanLedgerStateByKey = Object.create(null);
  let cleanLedgerUsed = false;
  // Sumber harus sama dengan Summary dan recovery. KM_UNIT_HISTORY yang masih
  // tersisa setelah migrasi bukan ledger tambahan untuk siklus yang sama.
  const historySheets = [getUnitKmHistorySheet_(ss)].filter(Boolean);
  const status3LSheet = ss.getSheetByName(UNIT_STATUS_3L_LOG_SHEET);

  const masterByPlate = Object.create(null);
  const currentUnits = [];
  const duplicateCounter = Object.create(null);
  const opOptions = Object.create(null);
  const tohOptions = Object.create(null);

  if (masterSheet && masterSheet.getLastRow() > 1 && masterSheet.getLastColumn() > 0) {
    const headers = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getDisplayValues()[0];
    const indexes = buildAliasIndex_(headers, MASTER_UNIT_FIELD_ALIASES);
    const rows = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, masterSheet.getLastColumn()).getValues();
    rows.forEach((row, offset) => {
      const nopol = indexes.nopol >= 0 ? normalizeDisplayNopol_(row[indexes.nopol]) : '';
      if (!nopol) return;
      const plateKey = normalizeNopol_(nopol);
      const operatingPoint = indexes.operatingPoint >= 0 ? displayValue_(row[indexes.operatingPoint]) : '';
      if (!canAccessOperatingPoint_(actor, operatingPoint)) return;
      duplicateCounter[plateKey] = (duplicateCounter[plateKey] || 0) + 1;
      const toh = indexes.toh >= 0 ? displayValue_(row[indexes.toh]) : '';
      if (operatingPoint) opOptions[operatingPoint] = true;
      if (toh) tohOptions[toh] = true;
      const unit = {
        rowNumber: offset + 2,
        nopol: nopol,
        operatingPoint: operatingPoint,
        toh: toh,
        unitType: indexes.unitType >= 0 ? displayValue_(row[indexes.unitType]) : '',
        masterStatus: indexes.masterStatus >= 0 ? displayValue_(row[indexes.masterStatus]) : 'ACTIVE',
        statusUnit: indexes.statusUnit >= 0 ? displayValue_(row[indexes.statusUnit]) : '',
        status3L: indexes.status3L >= 0 ? displayValue_(row[indexes.status3L]) : '',
        status3LUpdatedAt: indexes.status3LUpdatedAt >= 0 ? row[indexes.status3LUpdatedAt] : '',
        status3LUpdatedBy: indexes.status3LUpdatedBy >= 0 ? status3LActorDisplay_(displayValue_(row[indexes.status3LUpdatedBy])) : '',
        status3LNote: indexes.status3LNote >= 0 ? displayValue_(row[indexes.status3LNote]) : '',
        kmOdometer: indexes.kmOdometer >= 0 ? parseUnitKm_(row[indexes.kmOdometer]) : null,
        kmPhotoId: indexes.kmPhotoId >= 0 ? displayValue_(row[indexes.kmPhotoId]) : '',
        kmPhotoAddedAt: indexes.kmPhotoAddedAt >= 0 ? row[indexes.kmPhotoAddedAt] : '',
        kmPhotoAddedBy: indexes.kmPhotoAddedBy >= 0 ? displayValue_(row[indexes.kmPhotoAddedBy]) : '',
        kmDate: indexes.kmDate >= 0 ? row[indexes.kmDate] : '',
        kmUpdatedAt: indexes.kmUpdatedAt >= 0 ? row[indexes.kmUpdatedAt] : '',
        kmUpdatedBy: indexes.kmUpdatedBy >= 0 ? displayValue_(row[indexes.kmUpdatedBy]) : '',
        calibrationNumber: indexes.calibrationNumber >= 0 ? parseCalibrationNumber_(row[indexes.calibrationNumber]) : 0
      };
      if (normalizeLookupText_(unit.status3L) === 'DONE') {
        unit.status3LUpdatedBy = normalizeDoneStatus3LActor_(unit.status3L, unit.status3LUpdatedBy);
        unit.statusUnit = 'Ready to Depart';
      }
      masterByPlate[plateKey] = unit;
      if (opFilter && normalizeLookupText_(operatingPoint) !== opFilter) return;
      if (tohFilter && normalizeLookupText_(toh) !== tohFilter) return;
      currentUnits.push(unit);
    });
  }

  const submissionsByKey = Object.create(null);
  const finalsByKey = Object.create(null);
  const finalKeys = Object.create(null);
  // Semua pointer foto per siklus. Dipakai untuk membedakan FOTO ODOMETER ID
  // yang belum tersinkron di MASTER_UNIT dari bukti foto yang benar-benar hilang.
  const photoIdsByKey = Object.create(null);
  const currentFinalKeys = Object.create(null);
  // Resolver canonical v18 membedakan konflik legacy mentah dari pointer foto
  // yang benar-benar authoritative untuk siklus aktif.
  const historySources = historySheets.map(historySheet => analyticsReadHistory_(historySheet, UNIT_KM_LOG_HEADERS, diagnostics));
  const analyticsPhotoIndex = buildOdometerPhotoRecoveryIndex_(ss, historySources[0]);
  historySheets.forEach(historySheet => {
    const source = historySources[historySheets.indexOf(historySheet)];
    const map = source.map;
    const at = header => map[header] ? map[header] - 1 : -1;
    const rows = source.rows;
    rows.forEach(row => {
      const nopol = at('NOMOR POLISI') >= 0 ? normalizeDisplayNopol_(row[at('NOMOR POLISI')]) : '';
      const sequence = at('KALIBRASI KE') >= 0 ? parseCalibrationNumber_(row[at('KALIBRASI KE')]) : 0;
      const key = analyticsUnitKey_(nopol, sequence);
      const master = masterByPlate[normalizeNopol_(nopol)] || null;
      const recordType = at('RECORD TYPE') >= 0 ? normalizeLookupText_(row[at('RECORD TYPE')]) : '';
      const lifecycle = at('LIFECYCLE STATUS') >= 0 ? normalizeLookupText_(row[at('LIFECYCLE STATUS')]) : '';
      const rowPhotoId = historyOdometerPhotoId_(row, map);
      // Kualitas unit memakai coverage MASTER_UNIT saat ini. Transfer OP tidak
      // menghapus bukti historis siklus yang sama dari pemeriksaan foto/final.
      if (key && master && master.calibrationNumber === sequence) {
        if (rowPhotoId) {
          if (!photoIdsByKey[key]) photoIdsByKey[key] = Object.create(null);
          photoIdsByKey[key][rowPhotoId] = true;
        }
        if (recordType === 'FINAL_SNAPSHOT' && lifecycle === 'LOCKED') currentFinalKeys[key] = true;
      }
      const operatingPoint = (at('OPERATING POINT') >= 0 ? displayValue_(row[at('OPERATING POINT')]) : '') || (master ? master.operatingPoint : '');
      const toh = (at('NAMA TOH/WOH') >= 0 ? displayValue_(row[at('NAMA TOH/WOH')]) : '') || (master ? master.toh : '');
      if (!canAccessOperatingPoint_(actor, operatingPoint)) return;
      if (operatingPoint) opOptions[operatingPoint] = true;
      if (toh) tohOptions[toh] = true;
      if (opFilter && normalizeLookupText_(operatingPoint) !== opFilter) return;
      if (tohFilter && normalizeLookupText_(toh) !== tohFilter) return;
      if (!key) { if (nopol) diagnostics.missingSequenceRows++; return; }
      const timestamp = at('TIMESTAMP') >= 0 ? analyticsDateValue_(row[at('TIMESTAMP')]) : '';
      if (recordType === 'FINAL_SNAPSHOT' && lifecycle === 'LOCKED') {
        const completedRaw = at('COMPLETED AT') >= 0 ? row[at('COMPLETED AT')] : '';
        const completedAt = analyticsDateValue_(completedRaw) || timestamp;
        const completedBy = normalizeDoneStatus3LActor_('DONE',
          (at('COMPLETED BY') >= 0 ? displayValue_(row[at('COMPLETED BY')]) : '') ||
          (at('STATUS 3L UPDATED BY') >= 0 ? displayValue_(row[at('STATUS 3L UPDATED BY')]) : '') ||
          (at('USER EMAIL') >= 0 ? displayValue_(row[at('USER EMAIL')]) : ''));
        const dataAfter = at('DATA SETELAH') >= 0 ? safeParseHistoryJson_(row[at('DATA SETELAH')]) : {};
        const candidate = {
          key: key, nopol: nopol, calibrationNumber: sequence, completedAt: completedAt,
          completedBy: completedBy || KM_STATUS_3L_SHEET_FALLBACK, operatingPoint: operatingPoint, toh: toh,
          statusUnit: cleanText_(dataAfter.statusUnit) || (at('STATUS UNIT') >= 0 ? displayValue_(row[at('STATUS UNIT')]) : 'Ready to Depart')
        };
        if (!completedAt) { diagnostics.invalidDateRows++; return; }
        if (!finalsByKey[key] || !finalsByKey[key].completedAt || completedAt.getTime() < finalsByKey[key].completedAt.getTime()) {
          finalsByKey[key] = candidate;
        }
        finalKeys[key] = true;
        return;
      }
      if (!timestamp) { diagnostics.invalidDateRows++; return; }
      const candidate = {
        key: key, nopol: nopol, calibrationNumber: sequence, openedAt: timestamp,
        synthetic: at('REASON') >= 0 && normalizeLookupText_(row[at('REASON')]) === 'LEGACY_MASTER_SNAPSHOT',
        openedBy: at('USER EMAIL') >= 0 ? displayValue_(row[at('USER EMAIL')]) : '',
        operatingPoint: operatingPoint, toh: toh,
        statusUnitAtOpen: at('STATUS UNIT') >= 0 ? displayValue_(row[at('STATUS UNIT')]) : '',
        kmOdometer: at('KM ODOMETER') >= 0 ? parseUnitKm_(row[at('KM ODOMETER')]) : null,
        photoId: rowPhotoId
      };
      if (!submissionsByKey[key] || timestamp.getTime() < submissionsByKey[key].openedAt.getTime() ||
          (timestamp.getTime() === submissionsByKey[key].openedAt.getTime() && submissionsByKey[key].synthetic && !candidate.synthetic)) {
        submissionsByKey[key] = candidate;
      }
    });
  });

  // Primary analytics source v22: satu baris = satu pengajuan/siklus.
  // Raw history tetap dibaca di atas hanya sebagai evidence untuk Data Quality.
  const cleanLedgerRows = analyticsReadCleanLedger_(cleanLedgerSheet, diagnostics);
  if (cleanLedgerRows.length) {
    Object.keys(submissionsByKey).forEach(key => delete submissionsByKey[key]);
    Object.keys(finalsByKey).forEach(key => delete finalsByKey[key]);
    cleanLedgerRows.forEach(item => {
      if (!item.openedAt) { diagnostics.invalidDateRows++; return; }
      if (!canAccessOperatingPoint_(actor, item.operatingPoint)) return;
      if (opFilter && normalizeLookupText_(item.operatingPoint) !== opFilter) return;
      if (tohFilter && normalizeLookupText_(item.toh) !== tohFilter) return;
      if (item.operatingPoint) opOptions[item.operatingPoint] = true;
      if (item.toh) tohOptions[item.toh] = true;
      submissionsByKey[item.key] = {
        key: item.key,
        nopol: item.nopol,
        calibrationNumber: item.calibrationNumber,
        openedAt: item.openedAt,
        openedBy: item.openedBy,
        operatingPoint: item.operatingPoint,
        toh: item.toh,
        statusUnitAtOpen: '',
        kmOdometer: item.kmOdometer,
        photoId: item.photoId,
        requestId: item.requestId,
        historyId: item.historyId,
        synthetic: false
      };
      cleanLedgerStateByKey[item.key] = item;
      if (normalizeLookupText_(item.status3L) === 'DONE' && item.completedAt) {
        finalsByKey[item.key] = {
          key: item.key,
          nopol: item.nopol,
          calibrationNumber: item.calibrationNumber,
          completedAt: item.completedAt,
          completedBy: normalizeDoneStatus3LActor_('DONE', item.completedBy || item.status3LUpdatedBy),
          operatingPoint: item.operatingPoint,
          toh: item.toh,
          statusUnit: 'Ready to Depart',
          synthetic: false
        };
        finalKeys[item.key] = true;
      }
    });
    cleanLedgerUsed = true;
    diagnostics.primarySource = ANALYTICS_CALIBRATION_SHEET;
    diagnostics.cleanLedgerRows = cleanLedgerRows.length;
  } else {
    diagnostics.primarySource = UNIT_KM_LOG_SHEET;
    diagnostics.warnings.push('ANALYTICS_KALIBRASI_KM kosong/belum tersedia; analytics memakai fallback history. Jalankan Rebuild Analytics Ledger.');
  }

  // Legacy safety hanya dipakai bila clean ledger belum siap.
  if (!cleanLedgerUsed) Object.keys(masterByPlate).forEach(plateKey => {
    const unit = masterByPlate[plateKey];
    if (!canAccessOperatingPoint_(actor, unit.operatingPoint)) return;
    if (opFilter && normalizeLookupText_(unit.operatingPoint) !== opFilter) return;
    if (tohFilter && normalizeLookupText_(unit.toh) !== tohFilter) return;
    const key = analyticsUnitKey_(unit.nopol, unit.calibrationNumber);
    const openedAt = analyticsDateValue_(unit.kmUpdatedAt);
    if (key && openedAt && !submissionsByKey[key]) {
      diagnostics.recoveredSubmissions++;
      submissionsByKey[key] = {
        key: key, nopol: unit.nopol, calibrationNumber: unit.calibrationNumber,
        openedAt: openedAt, openedBy: unit.kmUpdatedBy, operatingPoint: unit.operatingPoint,
        toh: unit.toh, statusUnitAtOpen: unit.statusUnit, kmOdometer: unit.kmOdometer,
        photoId: unit.kmPhotoId, synthetic: true
      };
    }
  });

  const completionAudit = buildStatus3LAuditEvidenceMap_(status3LSheet);
  const latest3LByKey = Object.create(null);
  const allKendalaEvents = [];
  const seenStatusEvents = Object.create(null);
  if (status3LSheet && status3LSheet.getLastRow() > 1) {
    const source = analyticsReadHistory_(status3LSheet, UNIT_STATUS_3L_LOG_HEADERS, diagnostics);
    const map = source.map;
    const at = header => map[header] ? map[header] - 1 : -1;
    const rows = source.rows;
    rows.forEach(row => {
      const nopol = at('NOMOR POLISI') >= 0 ? normalizeDisplayNopol_(row[at('NOMOR POLISI')]) : '';
      const sequence = at('KALIBRASI KE') >= 0 ? parseCalibrationNumber_(row[at('KALIBRASI KE')]) : 0;
      const key = analyticsUnitKey_(nopol, sequence);
      const timestamp = at('TIMESTAMP') >= 0 ? analyticsDateValue_(row[at('TIMESTAMP')]) : '';
      const status3L = at('STATUS BARU') >= 0 ? displayValue_(row[at('STATUS BARU')]) : '';
      const user = at('USER EMAIL') >= 0 ? status3LActorDisplay_(displayValue_(row[at('USER EMAIL')])) : '';
      const note = at('CATATAN') >= 0 ? displayValue_(row[at('CATATAN')]) : '';
      const submission = submissionsByKey[key] || null;
      const master = masterByPlate[normalizeNopol_(nopol)] || null;
      const operatingPoint = submission ? submission.operatingPoint : (master ? master.operatingPoint : '');
      const toh = submission ? submission.toh : (master ? master.toh : '');
      if (!canAccessOperatingPoint_(actor, operatingPoint)) return;
      if (operatingPoint) opOptions[operatingPoint] = true;
      if (toh) tohOptions[toh] = true;
      if (opFilter && normalizeLookupText_(operatingPoint) !== opFilter) return;
      if (tohFilter && normalizeLookupText_(toh) !== tohFilter) return;
      if (!key) { if (nopol) diagnostics.missingSequenceRows++; return; }
      if (!timestamp) { diagnostics.invalidDateRows++; return; }
      const current = latest3LByKey[key];
      if (!current || timestamp.getTime() >= current.timestamp.getTime()) {
        latest3LByKey[key] = { status3L: status3L, timestamp: timestamp, updatedBy: user, note: note };
      }
      const noteOnly = at('REASON') >= 0 && /NOTE_UPDATED/.test(normalizeLookupText_(row[at('REASON')]));
      if (!cleanLedgerUsed && !noteOnly && normalizeLookupText_(status3L) === 'DONE' && !finalKeys[key] &&
          (!finalsByKey[key] || timestamp < finalsByKey[key].completedAt)) {
        finalsByKey[key] = {
          key: key, nopol: nopol, calibrationNumber: sequence, completedAt: timestamp,
          completedBy: user || KM_STATUS_3L_SHEET_FALLBACK, operatingPoint: operatingPoint, toh: toh,
          statusUnit: at('STATUS UNIT SESUDAH') >= 0 ? displayValue_(row[at('STATUS UNIT SESUDAH')]) : ''
        };
      }
      const eventId = at('UPDATE ID') >= 0 ? cleanText_(row[at('UPDATE ID')]) : '';
      const eventKey = eventId || [key, timestamp.getTime(), status3L, user, note].join('|');
      if (normalizeLookupText_(status3L) === 'KENDALA' && !seenStatusEvents[eventKey]) {
        seenStatusEvents[eventKey] = true;
        allKendalaEvents.push({ key: key, nopol: nopol, timestamp: timestamp, note: analyticsNormalizeNote_(note), operatingPoint: operatingPoint, toh: toh });
      }
    });
  }

  // Legacy DONE is recoverable only with an explicit completion timestamp.
  currentUnits.forEach(unit => {
    if (cleanLedgerUsed) return;
    const key = analyticsUnitKey_(unit.nopol, unit.calibrationNumber);
    const completedAt = analyticsDateValue_(unit.status3LUpdatedAt);
    if (key && completedAt && normalizeLookupText_(unit.status3L) === 'DONE' && !finalsByKey[key]) {
      finalsByKey[key] = { key: key, nopol: unit.nopol, calibrationNumber: unit.calibrationNumber,
        completedAt: completedAt, completedBy: unit.status3LUpdatedBy, operatingPoint: unit.operatingPoint,
        toh: unit.toh, statusUnit: unit.statusUnit, synthetic: true };
      diagnostics.recoveredCompletions++;
    }
  });
  const kendalaEvents = allKendalaEvents.filter(item => analyticsInRange_(item.timestamp, range));
  Object.keys(finalsByKey).forEach(key => {
    const item = finalsByKey[key];
    const evidence = status3LAuditEvidenceFor_(completionAudit, item.nopol, item.calibrationNumber, 'DONE');
    if (!isResolvedStatus3LActor_(item.completedBy) && evidence && evidence.resolvedActor) {
      item.completedBy = evidence.updatedBy;
      item.attributionSource = evidence.source;
    }
  });
  const allSubmissions = Object.keys(submissionsByKey).map(key => submissionsByKey[key]);
  const allFinals = Object.keys(finalsByKey).map(key => finalsByKey[key]);

  // Enrich every request, not only requests inside the selected chart period. This lets
  // backlog/repeat analytics use ANALYTICS_KALIBRASI_KM as their primary fact table too.
  allSubmissions.forEach(item => {
    const final = finalsByKey[item.key] || null;
    item.final = final;
    item.leadHours = final && !final.synthetic && !item.synthetic ? analyticsLeadHours_(item.openedAt, final.completedAt) : null;
    const current = masterByPlate[normalizeNopol_(item.nopol)] || null;
    const latest = latest3LByKey[item.key] || null;
    const ledgerState = cleanLedgerStateByKey[item.key] || null;
    item.currentStatus3L = ledgerState ? ledgerState.status3L : (final ? 'DONE' :
      (current && current.calibrationNumber === item.calibrationNumber ? current.status3L : (latest ? latest.status3L : '')));
    item.currentStatusUnit = ledgerState ? ledgerState.statusUnit : (final ? final.statusUnit :
      (current && current.calibrationNumber === item.calibrationNumber ? current.statusUnit : item.statusUnitAtOpen));
  });

  const periodSubmissions = allSubmissions.filter(item => analyticsInRange_(item.openedAt, range));
  const periodCompletions = allFinals.filter(item => analyticsInRange_(item.completedAt, range));

  // Backlog is the latest request per NOPOL. When clean ledger is available, MASTER_UNIT
  // is no longer the analytics source; it remains only an operational/data-quality source.
  const latestRequestByPlate = Object.create(null);
  allSubmissions.forEach(item => {
    const key = normalizeNopol_(item.nopol);
    if (!key) return;
    const existing = latestRequestByPlate[key];
    if (!existing || item.calibrationNumber > existing.calibrationNumber ||
        (item.calibrationNumber === existing.calibrationNumber && item.openedAt && existing.openedAt &&
          item.openedAt.getTime() > existing.openedAt.getTime())) latestRequestByPlate[key] = item;
  });
  const backlog = cleanLedgerUsed
    ? Object.keys(latestRequestByPlate).map(key => latestRequestByPlate[key]).filter(item =>
        normalizeLookupText_(item.currentStatus3L) !== 'DONE'
      ).map(item => ({
        nopol: item.nopol,
        calibrationNumber: item.calibrationNumber,
        operatingPoint: item.operatingPoint,
        toh: item.toh,
        status3L: item.currentStatus3L,
        statusUnit: item.currentStatusUnit,
        kmOdometer: item.kmOdometer,
        kmUpdatedAt: item.openedAt,
        ageHours: item.openedAt ? Math.max(0, (Date.now() - item.openedAt.getTime()) / 3600000) : null
      }))
    : currentUnits.filter(unit => {
        if (normalizeLookupText_(unit.masterStatus) === 'INACTIVE' || !unit.calibrationNumber) return false;
        return normalizeLookupText_(unit.status3L) !== 'DONE';
      }).map(unit => {
        const submission = submissionsByKey[analyticsUnitKey_(unit.nopol, unit.calibrationNumber)];
        const openedAt = submission ? submission.openedAt : analyticsDateValue_(unit.kmUpdatedAt);
        const ageHours = openedAt ? Math.max(0, (Date.now() - openedAt.getTime()) / 3600000) : null;
        return Object.assign({}, unit, { ageHours: ageHours });
      });

  const bucketKeys = analyticsBuildBucketKeys_(range, grain);
  const trendMap = Object.create(null);
  bucketKeys.forEach(key => { trendMap[key] = { key: key, submissions: 0, done: 0, kendala: 0 }; });
  periodSubmissions.forEach(item => {
    const key = analyticsBucketKey_(item.openedAt, grain);
    if (!trendMap[key]) trendMap[key] = { key: key, submissions: 0, done: 0, kendala: 0 };
    trendMap[key].submissions++;
  });
  periodCompletions.forEach(item => {
    const key = analyticsBucketKey_(item.completedAt, grain);
    if (!trendMap[key]) trendMap[key] = { key: key, submissions: 0, done: 0, kendala: 0 };
    trendMap[key].done++;
  });
  kendalaEvents.forEach(item => {
    const key = analyticsBucketKey_(item.timestamp, grain);
    if (!trendMap[key]) trendMap[key] = { key: key, submissions: 0, done: 0, kendala: 0 };
    trendMap[key].kendala++;
  });
  const trend = Object.keys(trendMap).sort().map(key => trendMap[key]);

  const pipelineCounts = { WAITING: 0, PROCESS: 0, KENDALA: 0, DONE: 0, OTHER: 0, READY: 0 };
  periodSubmissions.forEach(item => {
    const status = normalizeLookupText_(item.currentStatus3L);
    if (status === 'WAITING') pipelineCounts.WAITING++;
    else if (status === 'PROCESS') pipelineCounts.PROCESS++;
    else if (status === 'KENDALA') pipelineCounts.KENDALA++;
    else if (status === 'DONE') pipelineCounts.DONE++;
    else pipelineCounts.OTHER++;
    if (status === 'DONE' && normalizeLookupText_(item.currentStatusUnit) === normalizeLookupText_('Ready to Depart')) pipelineCounts.READY++;
  });

  const completedLeadHours = [];
  allFinals.forEach(item => {
    const submission = submissionsByKey[item.key];
    const lead = submission && !submission.synthetic && !item.synthetic ? analyticsLeadHours_(submission.openedAt, item.completedAt) : null;
    item.leadHours = lead;
    if (lead !== null && analyticsInRange_(item.completedAt, range)) completedLeadHours.push(lead);
  });
  const averageLeadHours = completedLeadHours.length ? completedLeadHours.reduce((a,b)=>a+b,0) / completedLeadHours.length : null;
  const slaMet = completedLeadHours.filter(value => value <= slaHours).length;

  const agingBands = [
    { key: 'LT2', label: '< 2 jam', count: 0 },
    { key: 'H2_4', label: '2–4 jam', count: 0 },
    { key: 'H4_8', label: '4–8 jam', count: 0 },
    { key: 'H8_24', label: '8–24 jam', count: 0 },
    { key: 'GT24', label: '≥ 24 jam', count: 0 },
    { key: 'UNKNOWN', label: 'Tanggal belum ada', count: 0 }
  ];
  backlog.forEach(item => {
    const age = item.ageHours;
    if (age === null) agingBands[5].count++;
    else if (age >= 24) agingBands[4].count++;
    else if (age < 2) agingBands[0].count++;
    else if (age < 4) agingBands[1].count++;
    else if (age < 8) agingBands[2].count++;
    else agingBands[3].count++;
  });

  function aggregateDimension(field) {
    const groups = Object.create(null);
    periodSubmissions.forEach(item => {
      const name = cleanText_(item[field]) || 'Belum ditentukan';
      if (!groups[name]) groups[name] = { name: name, submissions: 0, done: 0, leadTotal: 0, leadCount: 0, backlog: 0 };
      groups[name].submissions++;
      if (item.final) {
        groups[name].done++;
        if (item.leadHours !== null) { groups[name].leadTotal += item.leadHours; groups[name].leadCount++; }
      }
    });
    backlog.forEach(item => {
      const name = cleanText_(item[field]) || 'Belum ditentukan';
      if (!groups[name]) groups[name] = { name: name, submissions: 0, done: 0, leadTotal: 0, leadCount: 0, backlog: 0 };
      groups[name].backlog++;
    });
    return Object.keys(groups).map(name => {
      const item = groups[name];
      return {
        name: item.name, submissions: item.submissions, done: item.done, backlog: item.backlog,
        completionRate: analyticsPercent_(item.done, item.submissions),
        avgLeadHours: item.leadCount ? analyticsRound_(item.leadTotal / item.leadCount, 1) : 0
      };
    }).sort((a,b) => (b.submissions - a.submissions) || (b.backlog - a.backlog) || a.name.localeCompare(b.name));
  }

  // Coverage kalibrasi adalah snapshot populasi MASTER_UNIT, bukan jumlah event
  // pada periode. Satu Nopol hanya dihitung sekali dan unit INACTIVE dikeluarkan.
  // Definisi ini konsisten dengan dashboard operasional: Kalibrasi Ke > 0 berarti
  // unit sudah memiliki minimal satu siklus/pengajuan kalibrasi yang tercatat.
  const coverageUnitByPlate = Object.create(null);
  currentUnits.forEach(unit => {
    const plateKey = normalizeNopol_(unit.nopol);
    if (!plateKey || normalizeLookupText_(unit.masterStatus) === 'INACTIVE') return;
    coverageUnitByPlate[plateKey] = unit;
  });
  const coverageUnits = Object.keys(coverageUnitByPlate).map(key => coverageUnitByPlate[key]);

  function calibrationCoverageSummary(units) {
    const source = Array.isArray(units) ? units : [];
    const calibratedUnits = source.filter(unit => Number(unit.calibrationNumber) > 0).length;
    const totalUnits = source.length;
    const uncalibratedUnits = Math.max(0, totalUnits - calibratedUnits);
    return {
      totalUnits: totalUnits,
      calibratedUnits: calibratedUnits,
      uncalibratedUnits: uncalibratedUnits,
      calibratedPercent: analyticsPercent_(calibratedUnits, totalUnits),
      uncalibratedPercent: analyticsPercent_(uncalibratedUnits, totalUnits)
    };
  }

  function aggregateCalibrationCoverage(field) {
    const groups = Object.create(null);
    coverageUnits.forEach(unit => {
      const name = cleanText_(unit[field]) || 'Belum dipetakan';
      if (!groups[name]) groups[name] = [];
      groups[name].push(unit);
    });
    return Object.keys(groups).map(name => Object.assign(
      { name: name }, calibrationCoverageSummary(groups[name])
    )).sort((a, b) =>
      (a.calibratedPercent - b.calibratedPercent) ||
      (b.uncalibratedUnits - a.uncalibratedUnits) ||
      a.name.localeCompare(b.name)
    );
  }

  const calibrationCoverage = Object.assign(calibrationCoverageSummary(coverageUnits), {
    operatingPointCount: Object.keys(coverageUnits.reduce((acc, unit) => {
      const operatingPoint = cleanText_(unit.operatingPoint);
      if (operatingPoint) acc[operatingPoint] = true;
      return acc;
    }, Object.create(null))).length,
    unmappedUnits: coverageUnits.filter(unit => !cleanText_(unit.operatingPoint)).length,
    definition: 'CALIBRATION_NUMBER_GT_ZERO',
    population: 'ACTIVE_UNIQUE_UNITS'
  });

  const activeUnitByPlate = Object.create(null);
  coverageUnits.forEach(unit => { activeUnitByPlate[normalizeNopol_(unit.nopol)] = unit; });
  const activePeriodSubmissions = periodSubmissions.filter(item => activeUnitByPlate[normalizeNopol_(item.nopol)]);

  const periodSubmittedByPlate = Object.create(null);
  activePeriodSubmissions.forEach(item => {
    const plateKey = normalizeNopol_(item.nopol);
    if (plateKey && activeUnitByPlate[plateKey]) periodSubmittedByPlate[plateKey] = true;
  });
  const submittedUnitsInRange = Object.keys(periodSubmittedByPlate).length;
  const submissionCoverage = {
    totalUnits: coverageUnits.length,
    submittedUnits: submittedUnitsInRange,
    unsubmittedUnits: Math.max(0, coverageUnits.length - submittedUnitsInRange),
    submittedPercent: analyticsPercent_(submittedUnitsInRange, coverageUnits.length),
    unsubmittedPercent: analyticsPercent_(Math.max(0, coverageUnits.length - submittedUnitsInRange), coverageUnits.length),
    submissionEvents: activePeriodSubmissions.length,
    definition: 'DISTINCT_ACTIVE_UNIT_WITH_SUBMISSION_IN_SELECTED_RANGE',
    population: 'CURRENT_ACTIVE_UNIQUE_UNITS'
  };

  function aggregateSubmissionCoverage(field) {
    const groups = Object.create(null);
    coverageUnits.forEach(unit => {
      const name = cleanText_(unit[field]) || 'Belum dipetakan';
      if (!groups[name]) groups[name] = [];
      groups[name].push(unit);
    });
    return Object.keys(groups).map(name => {
      const units = groups[name];
      const submittedUnits = units.filter(unit => periodSubmittedByPlate[normalizeNopol_(unit.nopol)]).length;
      const totalUnits = units.length;
      return {
        name: name,
        totalUnits: totalUnits,
        submittedUnits: submittedUnits,
        unsubmittedUnits: Math.max(0, totalUnits - submittedUnits),
        submittedPercent: analyticsPercent_(submittedUnits, totalUnits),
        unsubmittedPercent: analyticsPercent_(Math.max(0, totalUnits - submittedUnits), totalUnits)
      };
    }).sort((a, b) =>
      (a.submittedPercent - b.submittedPercent) ||
      (b.unsubmittedUnits - a.unsubmittedUnits) ||
      a.name.localeCompare(b.name)
    );
  }

  const bucketSubmittedUnits = Object.create(null);
  trend.forEach(item => { item.submissions = 0; });
  activePeriodSubmissions.forEach(item => {
    const plateKey = normalizeNopol_(item.nopol);
    if (!plateKey || !activeUnitByPlate[plateKey]) return;
    const bucketKey = analyticsBucketKey_(item.openedAt, grain);
    if (trendMap[bucketKey]) trendMap[bucketKey].submissions++;
    if (!bucketSubmittedUnits[bucketKey]) bucketSubmittedUnits[bucketKey] = Object.create(null);
    if (!bucketSubmittedUnits[bucketKey][plateKey]) {
      bucketSubmittedUnits[bucketKey][plateKey] = {
        nopol: activeUnitByPlate[plateKey].nopol,
        submissionCount: 0,
        latestSubmittedAt: '',
        calibrationNumbers: []
      };
    }
    const detail = bucketSubmittedUnits[bucketKey][plateKey];
    detail.submissionCount++;
    const openedIso = item.openedAt ? item.openedAt.toISOString() : '';
    if (!detail.latestSubmittedAt || openedIso > detail.latestSubmittedAt) detail.latestSubmittedAt = openedIso;
    if (item.calibrationNumber && !detail.calibrationNumbers.includes(item.calibrationNumber)) {
      detail.calibrationNumbers.push(item.calibrationNumber);
    }
  });
  trend.forEach(item => {
    const bucket = bucketSubmittedUnits[item.key] || Object.create(null);
    const details = Object.keys(bucket).map(key => bucket[key]).sort((a, b) => a.nopol.localeCompare(b.nopol));
    details.forEach(detail => detail.calibrationNumbers.sort((a, b) => a - b));
    item.totalUnits = coverageUnits.length;
    item.submittedUnits = details.length;
    item.unsubmittedUnits = Math.max(0, coverageUnits.length - details.length);
    item.submittedPercent = analyticsPercent_(details.length, coverageUnits.length);
    item.unsubmittedPercent = analyticsPercent_(Math.max(0, coverageUnits.length - details.length), coverageUnits.length);
    item.submittedUnitDetails = details;
  });

  const kendalaReasonMap = Object.create(null);
  kendalaEvents.forEach(item => {
    const name = item.note || 'Tanpa catatan';
    if (!kendalaReasonMap[name]) kendalaReasonMap[name] = { name: name, count: 0, units: Object.create(null) };
    kendalaReasonMap[name].count++;
    kendalaReasonMap[name].units[normalizeNopol_(item.nopol)] = true;
  });
  const kendalaReasons = Object.keys(kendalaReasonMap).map(name => ({
    name: name, count: kendalaReasonMap[name].count,
    units: Object.keys(kendalaReasonMap[name].units).length
  })).sort((a,b) => b.count - a.count).slice(0, 10);

  const repeatSubmissions = periodSubmissions.filter(item => item.calibrationNumber > 1);
  const repeatDistribution = [
    { label: 'Kalibrasi ke-1', count: periodSubmissions.filter(item => item.calibrationNumber === 1).length },
    { label: 'Kalibrasi ke-2', count: periodSubmissions.filter(item => item.calibrationNumber === 2).length },
    { label: 'Kalibrasi ke-3', count: periodSubmissions.filter(item => item.calibrationNumber === 3).length },
    { label: 'Kalibrasi ke-4+', count: periodSubmissions.filter(item => item.calibrationNumber >= 4).length }
  ];
  const repeatTopSource = cleanLedgerUsed
    ? Object.keys(latestRequestByPlate).map(key => latestRequestByPlate[key]).map(item => ({
        nopol: item.nopol,
        calibrationNumber: item.calibrationNumber,
        operatingPoint: item.operatingPoint,
        toh: item.toh,
        status3L: item.currentStatus3L,
        statusUnit: item.currentStatusUnit
      }))
    : currentUnits;
  const repeatTop = repeatTopSource.filter(item => item.calibrationNumber > 1)
    .sort((a,b) => b.calibrationNumber - a.calibrationNumber)
    .slice(0, 12).map(item => ({
      nopol: item.nopol, calibrationNumber: item.calibrationNumber,
      operatingPoint: item.operatingPoint, toh: item.toh,
      status3L: item.status3L, statusUnit: item.statusUnit
    }));

  const teamMap = Object.create(null);
  periodCompletions.forEach(item => {
    const email = normalizeEmail_(item.completedBy) || cleanText_(item.completedBy) || KM_STATUS_3L_SHEET_FALLBACK;
    if (!teamMap[email]) teamMap[email] = { user: email, done: 0, leadTotal: 0, leadCount: 0, slaMet: 0, ops: Object.create(null) };
    const team = teamMap[email];
    team.done++;
    if (item.operatingPoint) team.ops[item.operatingPoint] = true;
    if (item.leadHours !== null && item.leadHours !== undefined) {
      team.leadTotal += item.leadHours; team.leadCount++;
      if (item.leadHours <= slaHours) team.slaMet++;
    }
  });
  const teamWorkload = Object.keys(teamMap).map(key => {
    const item = teamMap[key];
    return {
      user: item.user, done: item.done,
      avgLeadHours: item.leadCount ? analyticsRound_(item.leadTotal / item.leadCount, 1) : null,
      slaAchievement: item.leadCount ? analyticsPercent_(item.slaMet, item.leadCount) : null,
      operatingPoints: Object.keys(item.ops).length
    };
  }).sort((a,b) => b.done - a.done).slice(0, 20);

  const issueDefs = {
    doneNotReady: { label: 'DONE tetapi belum Ready to Depart', count: 0, samples: [] },
    readyNotDone: { label: 'Ready to Depart tetapi Status 3L bukan DONE', count: 0, samples: [] },
    missingOdometer: { label: 'Kalibrasi tanpa KM Odometer', count: 0, samples: [] },
    photoIdNotSynced: { label: 'ID Foto tersedia di history tetapi belum tersinkron ke MASTER_UNIT', count: 0, samples: [] },
    missingPhoto: { label: 'Kalibrasi benar-benar tanpa bukti Foto Odometer', count: 0, samples: [] },
    photoIdAmbiguous: { label: 'Konflik ID Foto belum memiliki pointer canonical', count: 0, samples: [] },
    photoIdMismatch: { label: 'Pointer Foto MASTER_UNIT berbeda dengan canonical history', count: 0, samples: [] },
    photoMetadata: { label: 'Waktu / uploader foto tidak sesuai dengan ID foto di history', count: 0, samples: [] },
    doneMissingFinal: { label: 'DONE tanpa FINAL_SNAPSHOT / LOCKED', count: 0, samples: [] },
    doneUpdaterUnresolved: { label: 'DONE dengan updater 3L belum teridentifikasi', count: 0, samples: [] },
    incompleteCoverage: { label: 'Operating Point / TOH-WOH belum lengkap', count: 0, samples: [] },
    duplicateNopol: { label: 'Duplicate Nomor Polisi', count: 0, samples: [] }
  };
  const affected = Object.create(null);
  function addIssue(name, unit, meta) {
    const issue = issueDefs[name]; if (!issue) return;
    issue.count++;
    affected[unit.rowNumber] = true;
    if (issue.samples.length < 12) issue.samples.push(Object.assign({
      nopol: unit.nopol, operatingPoint: unit.operatingPoint,
      status3L: unit.status3L, statusUnit: unit.statusUnit
    }, meta || {}));
  }
  currentUnits.forEach(unit => {
    const done = normalizeLookupText_(unit.status3L) === 'DONE';
    const ready = normalizeLookupText_(unit.statusUnit) === normalizeLookupText_('Ready to Depart');
    if (done && !ready) addIssue('doneNotReady', unit);
    if (ready && !done) addIssue('readyNotDone', unit);
    if (unit.calibrationNumber > 0 && unit.kmOdometer === null) addIssue('missingOdometer', unit);
    if (unit.calibrationNumber > 0) {
      const photoKey = analyticsUnitKey_(unit.nopol, unit.calibrationNumber);
      const masterPhotoId = normalizeOdometerPhotoId_(unit.kmPhotoId);
      const historyPhotoIds = photoIdsByKey[photoKey] ? Object.keys(photoIdsByKey[photoKey]) : [];
      const resolution = resolveOdometerPhotoRecoveryCandidate_(
        analyticsPhotoIndex, unit.nopol, unit.calibrationNumber, unit.kmDate, unit.kmUpdatedBy
      );
      const canonicalPhotoId = resolution.status === 'RECOVERABLE'
        ? normalizeOdometerPhotoId_(resolution.candidate && resolution.candidate.photoId) : '';
      if (!masterPhotoId) {
        if (resolution.status === 'RECOVERABLE' && canonicalPhotoId) {
          addIssue('photoIdNotSynced', unit, {
            photoId: canonicalPhotoId, calibrationNumber: unit.calibrationNumber,
            photoSource: analyticsPhotoIndex.historySheet || UNIT_KM_LOG_SHEET,
            photoResolution: resolution.match
          });
        } else if (resolution.status === 'AMBIGUOUS') {
          addIssue('photoIdAmbiguous', unit, {
            calibrationNumber: unit.calibrationNumber,
            historyPhotoIds: (resolution.candidates || []).map(item => item.photoId),
            photoCandidateCount: (resolution.candidates || []).length,
            photoSource: analyticsPhotoIndex.historySheet || UNIT_KM_LOG_SHEET
          });
        } else {
          addIssue('missingPhoto', unit);
        }
      } else if (resolution.status === 'RECOVERABLE' && canonicalPhotoId && canonicalPhotoId !== masterPhotoId) {
        addIssue('photoIdMismatch', unit, {
          masterPhotoId: masterPhotoId,
          canonicalPhotoId: canonicalPhotoId,
          historyPhotoIds: historyPhotoIds,
          calibrationNumber: unit.calibrationNumber,
          photoSource: analyticsPhotoIndex.historySheet || UNIT_KM_LOG_SHEET,
          photoResolution: resolution.match,
          photoCandidateCount: historyPhotoIds.length
        });
      } else if (resolution.status === 'AMBIGUOUS') {
        const candidates = (resolution.candidates || []).map(item => item.photoId);
        // Bila MASTER_UNIT memiliki pointer aktif tetapi history belum mempunyai
        // canonical resolution, issue tetap ditampilkan sampai admin menjalankan
        // Repair Photo Integrity. Repair tidak menghapus histori lama.
        addIssue('photoIdAmbiguous', unit, {
          masterPhotoId: masterPhotoId,
          historyPhotoIds: candidates,
          calibrationNumber: unit.calibrationNumber,
          photoCandidateCount: candidates.length,
          photoSource: analyticsPhotoIndex.historySheet || UNIT_KM_LOG_SHEET
        });
      }
      if (masterPhotoId) {
        const metadataCandidate = odometerPhotoCandidateForId_(
          analyticsPhotoIndex, unit.nopol, unit.calibrationNumber, masterPhotoId, unit.kmDate
        );
        const metadata = compareOdometerPhotoMetadata_(unit.kmPhotoAddedAt, unit.kmPhotoAddedBy, metadataCandidate);
        const metadataIssues = [];
        if (metadata.missingAt) metadataIssues.push('FOTO DITAMBAHKAN AT kosong');
        else if (metadata.invalidAt) metadataIssues.push('FOTO DITAMBAHKAN AT invalid');
        else if (metadata.futureAt) metadataIssues.push('FOTO DITAMBAHKAN AT future');
        else if (metadata.mismatchAt) metadataIssues.push('FOTO DITAMBAHKAN AT mismatch');
        if (metadata.missingBy) metadataIssues.push('FOTO DIUNGGAH OLEH kosong');
        else if (metadata.mismatchBy) metadataIssues.push('FOTO DIUNGGAH OLEH mismatch');
        if (metadataIssues.length) addIssue('photoMetadata', unit, {
          photoId: masterPhotoId,
          calibrationNumber: unit.calibrationNumber,
          issues: metadataIssues,
          historyEvidence: Boolean(metadataCandidate)
        });
      }
    }
    if (done && !currentFinalKeys[analyticsUnitKey_(unit.nopol, unit.calibrationNumber)]) addIssue('doneMissingFinal', unit);
    const updaterKey = normalizeLookupText_(unit.status3LUpdatedBy);
    if (done && !isResolvedStatus3LActor_(unit.status3LUpdatedBy)) addIssue('doneUpdaterUnresolved', unit);
    if (!cleanText_(unit.operatingPoint) || !cleanText_(unit.toh)) addIssue('incompleteCoverage', unit);
    if ((duplicateCounter[normalizeNopol_(unit.nopol)] || 0) > 1) addIssue('duplicateNopol', unit);
  });
  const dataQualityIssues = Object.keys(issueDefs).map(key => Object.assign({ key: key }, issueDefs[key]));
  const qualityScore = currentUnits.length ? Math.max(0, analyticsRound_(100 - (Object.keys(affected).length * 100 / currentUnits.length), 1)) : 100;

  const previousRange = analyticsPreviousRange_(range);
  const currentMetrics = analyticsSubmissionPeriodMetrics_(allSubmissions, range, activeUnitByPlate);
  const previousMetrics = analyticsSubmissionPeriodMetrics_(allSubmissions, previousRange, activeUnitByPlate);
  const observedDates = allSubmissions.map(item => item.openedAt)
    .concat(allFinals.map(item => item.completedAt), allKendalaEvents.map(item => item.timestamp)).filter(Boolean);
  const firstObserved = observedDates.length ? new Date(observedDates.reduce((min,value) => Math.min(min,value.getTime()), Infinity)) : null;
  const lastObserved = observedDates.length ? new Date(observedDates.reduce((max,value) => Math.max(max,value.getTime()), -Infinity)) : null;
  diagnostics.firstObservedDate = firstObserved ? dateOnlyKey_(firstObserved) : '';
  diagnostics.lastObservedDate = lastObserved ? dateOnlyKey_(lastObserved) : '';
  diagnostics.slaExcludedSamples = periodCompletions.length - completedLeadHours.length;
  const comparison = analyticsComparison_(currentMetrics, previousMetrics, previousRange);
  comparison.historyUnavailable = !firstObserved || previousRange.end < analyticsStartOfDay_(firstObserved);
  comparison.partialHistory = Boolean(firstObserved && previousRange.start < analyticsStartOfDay_(firstObserved));
  if (comparison.historyUnavailable) comparison.metrics.forEach(metric => {
    metric.previous = null; metric.delta = null; metric.percentChange = null;
  });
  const previousBuckets = analyticsBuildBucketKeys_({
    start: analyticsStartOfDay_(analyticsBucketKey_(range.start, grain) + (grain === 'MONTHLY' ? '-01' : '')),
    end: range.end
  }, grain);
  const comparisonRows = trend.map((item, index) => {
    const previous = index ? trend[index - 1] : null;
    return Object.assign({}, item, { netFlow: item.submissions - item.done,
      submissionDelta: previous ? item.submissions - previous.submissions : null,
      doneDelta: previous ? item.done - previous.done : null,
      submittedUnitDelta: previous ? item.submittedUnits - previous.submittedUnits : null,
      submittedPercentDelta: previous ? analyticsRound_(item.submittedPercent - previous.submittedPercent, 1) : null,
      partial: (index === trend.length - 1 && range.end.getTime() >= Date.now()) || (index === 0 && range.startKey !== previousBuckets[0] && grain !== 'MONTHLY') ||
        (grain === 'MONTHLY' && ((index === 0 && range.start.getDate() !== 1) ||
          (index === trend.length - 1 && new Date(range.end.getFullYear(), range.end.getMonth() + 1, 0).getDate() !== range.end.getDate()))) ||
        (grain === 'WEEKLY' && index === trend.length - 1 && range.end.getDay() !== 0)
    });
  });
  const result = {
    generatedAt: new Date().toISOString(),
    diagnostics: diagnostics,
    comparison: comparison,
    comparisonRows: comparisonRows,
    filters: {
      startDate: range.startKey, endDate: range.endKey, grain: grain,
      operatingPoint: cleanText_(input.operatingPoint), toh: cleanText_(input.toh), slaHours: slaHours
    },
    filterOptions: {
      operatingPoints: Object.keys(opOptions).sort(),
      toh: Object.keys(tohOptions).sort()
    },
    kpi: {
      submissions: periodSubmissions.length,
      done: periodCompletions.length,
      cohortDone: periodSubmissions.filter(item => Boolean(item.final)).length,
      completionRate: analyticsPercent_(periodSubmissions.filter(item => Boolean(item.final)).length, periodSubmissions.length),
      completionsInPeriod: periodCompletions.length,
      averageLeadHours: averageLeadHours == null ? null : analyticsRound_(averageLeadHours, 1),
      slaAchievement: completedLeadHours.length ? analyticsPercent_(slaMet, completedLeadHours.length) : null,
      activeBacklog: backlog.length,
      repeatRate: analyticsPercent_(repeatSubmissions.length, periodSubmissions.length),
      repeatSubmissions: repeatSubmissions.length,
      qualityScore: qualityScore,
      kendalaEvents: kendalaEvents.length
    },
    calibrationCoverage: calibrationCoverage,
    byOperatingPointCoverage: aggregateCalibrationCoverage('operatingPoint'),
    submissionCoverage: submissionCoverage,
    byOperatingPointSubmissionCoverage: aggregateSubmissionCoverage('operatingPoint'),
    activeUnits: coverageUnits.map(unit => ({
      nopol: unit.nopol,
      operatingPoint: unit.operatingPoint,
      toh: unit.toh,
      unitType: unit.unitType,
      statusUnit: unit.statusUnit,
      status3L: unit.status3L,
      calibrationNumber: unit.calibrationNumber
    })).sort((a, b) =>
      (a.operatingPoint || '').localeCompare(b.operatingPoint || '') || a.nopol.localeCompare(b.nopol)
    ),
    trend: trend,
    pipeline: pipelineCounts,
    sla: {
      targetHours: slaHours,
      completedSamples: completedLeadHours.length,
      met: slaMet,
      breached: Math.max(0, completedLeadHours.length - slaMet),
      achievement: completedLeadHours.length ? analyticsPercent_(slaMet, completedLeadHours.length) : null,
      excludedSamples: periodCompletions.length - completedLeadHours.length,
      averageLeadHours: averageLeadHours == null ? null : analyticsRound_(averageLeadHours, 1)
    },
    aging: {
      total: backlog.length,
      bands: agingBands,
      critical: backlog.filter(item => item.ageHours !== null && item.ageHours >= 24).length,
      unknownAge: backlog.filter(item => item.ageHours === null).length,
      overdueSla: backlog.filter(item => item.ageHours !== null && item.ageHours > slaHours).length,
      top: backlog.slice().sort((a,b) => (b.ageHours == null ? -1 : b.ageHours) - (a.ageHours == null ? -1 : a.ageHours)).slice(0, 15).map(item => ({
        nopol: item.nopol, operatingPoint: item.operatingPoint, toh: item.toh,
        status3L: item.status3L || 'Belum Status', ageHours: item.ageHours === null ? null : analyticsRound_(item.ageHours, 1),
        kmUpdatedAt: analyticsDateValue_(item.kmUpdatedAt) ? analyticsDateValue_(item.kmUpdatedAt).toISOString() : ''
      }))
    },
    byOperatingPoint: aggregateDimension('operatingPoint').slice(0, 20),
    byToh: aggregateDimension('toh').slice(0, 20),
    kendala: {
      totalEvents: kendalaEvents.length,
      uniqueUnits: Object.keys(kendalaEvents.reduce((acc,item)=>{acc[normalizeNopol_(item.nopol)]=true;return acc;},Object.create(null))).length,
      reasons: kendalaReasons
    },
    repeatCalibration: {
      repeatRate: analyticsPercent_(repeatSubmissions.length, periodSubmissions.length),
      repeatSubmissions: repeatSubmissions.length,
      distribution: repeatDistribution,
      topUnits: repeatTop
    },
    teamWorkload: teamWorkload,
    dataQuality: {
      score: qualityScore,
      affectedUnits: Object.keys(affected).length,
      totalUnits: currentUnits.length,
      issues: dataQualityIssues,
      photoIntegrity: {
        recoverable: issueDefs.photoIdNotSynced.count,
        trueMissing: issueDefs.missingPhoto.count,
        ambiguous: issueDefs.photoIdAmbiguous.count,
        mismatch: issueDefs.photoIdMismatch.count,
        metadata: issueDefs.photoMetadata.count,
        historyMetadata: analyticsPhotoIndex.historyMetadata || {},
        rawHistoryConflicts: Number(analyticsPhotoIndex.rawConflictCount || 0)
      }
    }
  };
  putCachedJson_(cacheKey, result, 60);
  return result;
}

function getMasterUnitSummaryData(nopol, sessionToken) {
  const actor = requirePermission_('VIEW_MASTER_UNIT', '', { token: sessionToken });
  const masterTarget = readMasterUnitSummaryTarget_(nopol);
  if (!masterTarget) throw new Error('Nomor polisi tidak ditemukan.');
  if (masterTarget.indexes.operatingPoint < 0) {
    throw new Error('Header Operating Point belum tersedia pada MASTER_UNIT.');
  }
  const operatingPoint = displayValue_(masterTarget.row[masterTarget.indexes.operatingPoint]);
  if (!canAccessOperatingPoint_(actor, operatingPoint)) {
    throw new Error('Akses ditolak: unit berada di luar scope Operating Point akun Anda.');
  }

  // Cache singkat berbasis revision. invalidateAppCaches_ mengganti revision setiap
  // ada perubahan workflow/master, sehingga response lama otomatis tidak dipakai.
  let revision = getCachedJson_('analytics-revision');
  if (!revision) {
    revision = Utilities.getUuid();
    putCachedJson_('analytics-revision', revision, 600);
  }
  const cacheKey = 'unit-summary-v2:' + sha256Hex_(JSON.stringify([
    revision, normalizeNopol_(nopol)
  ]));
  const cached = getCachedJson_(cacheKey);
  if (cached) return cached;

  const currentState = getMasterUnitCurrentSummaryState_(nopol, masterTarget);
  const rawStatus3LHistory = getMasterUnitStatus3LHistory_(nopol);
  const rawStatusHistory = getMasterUnitStatusHistory_(nopol);
  const status3LResult = reconcileSummaryStatus3LHistory_(rawStatus3LHistory, currentState);
  const statusUnitResult = reconcileSummaryStatusUnitHistory_(rawStatusHistory, currentState);
  const kmHistory = getMasterUnitKmHistory_(nopol, status3LResult.rows, currentState);
  const currentSequence = currentState ? parseCalibrationNumber_(currentState.calibrationNumber) : 0;
  const currentKmHistory = currentSequence
    ? kmHistory.find(item => parseCalibrationNumber_(item.calibrationNumber) === currentSequence) || null
    : null;
  const finalSnapshotMissing = Boolean(currentKmHistory && currentKmHistory.finalSnapshotMissing);
  const result = {
    fields: getMasterUnitDetail_(nopol, masterTarget),
    kmHistory: kmHistory,
    statusHistory: statusUnitResult.rows,
    status3LHistory: status3LResult.rows,
    currentState: currentState,
    integrity: {
      reconciled: Boolean(status3LResult.reconciled || statusUnitResult.reconciled || finalSnapshotMissing),
      status3LReconciled: Boolean(status3LResult.reconciled),
      statusUnitReconciled: Boolean(statusUnitResult.reconciled),
      finalSnapshotMissing: finalSnapshotMissing,
      sourceOfTruth: 'MASTER_UNIT',
      message: (status3LResult.reconciled || statusUnitResult.reconciled || finalSnapshotMissing)
        ? 'Ringkasan diselaraskan dengan status aktual MASTER_UNIT. Riwayat audit asli tetap dipertahankan.'
        : 'Ringkasan dan histori konsisten dengan status aktual.'
    }
  };
  putCachedJson_(cacheKey, result, 45);
  return result;
}