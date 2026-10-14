/**
 * Offer Importer — Google Apps Script edition.
 *
 * Browser side still does all the Excel work (identical output to the
 * GitHub Pages tool). The server only archives the uploaded workbook as a
 * Google Sheet and records a History row.
 */

var VERSION = '2.09.14.10.2026';
var FOLDER_ID = '1JKUqS-loYgZjvJULYY4xQmGnoIuVMrNt';
var LOG_SPREADSHEET_ID = '1Iw4_A5SnVXxgxkTVrVkkcUWGOMyqMo0cSflhEvnumkw';
var LOG_HEADERS = ['Timestamp', 'User_Email', 'File_Type', 'File_Name', 'Sheet_URL', 'Outputs', 'Hotel_Names'];
var MAX_HOTEL_NAMES = 60;
var FILE_TYPE_3PP = '3PP form';
var FILE_TYPE_UPDATE = 'update form';
/** Only these people can reveal the raw log spreadsheet link. */
var LOG_VIEWERS = ['b.gheorghe@secretescapes.com'];
var MAX_UPLOAD_BASE64 = 25 * 1024 * 1024;

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Offer Importer ' + VERSION)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function getBootstrap() {
  return { version: VERSION };
}

/** Revealed by tapping the version pill three times. null for anyone else. */
function getLogLink() {
  var email = String(getUserEmail_()).trim().toLowerCase();
  if (LOG_VIEWERS.indexOf(email) === -1) return null;
  return { url: 'https://docs.google.com/spreadsheets/d/' + LOG_SPREADSHEET_ID + '/edit' };
}

function getUserEmail_() {
  var email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (ignored) {}
  if (!email) {
    try { email = Session.getEffectiveUser().getEmail(); } catch (ignored) {}
  }
  return email || 'unknown';
}

/* ------------------------------------------------------------------ *
 * Upload: archive as Google Sheet + write the History row
 * ------------------------------------------------------------------ */

/**
 * payload: { fileName, base64, mimeType, sheetNames: [], hotelNames: [], outputCount }
 */
function logUpload(payload) {
  payload = payload || {};
  var fileName = String(payload.fileName || 'workbook');
  var base64 = String(payload.base64 || '');

  if (!base64) throw new Error('No file content was received.');
  if (base64.length > MAX_UPLOAD_BASE64) {
    throw new Error('That workbook is too large to archive (over 25 MB). The upload files were still generated.');
  }

  var fileType = detectFileType_(payload.sheetNames);
  var blob = Utilities.newBlob(
    Utilities.base64Decode(base64),
    payload.mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    fileName
  );

  var created = Drive.Files.create(
    {
      name: stripExtension_(fileName),
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [FOLDER_ID]
    },
    blob,
    { supportsAllDrives: true }
  );

  var url = 'https://docs.google.com/spreadsheets/d/' + created.id + '/edit';
  appendLogRow_([
    new Date(),
    getUserEmail_(),
    fileType,
    fileName,
    url,
    Number(payload.outputCount || 0),
    joinHotelNames_(payload.hotelNames)
  ]);

  return { fileType: fileType, url: url, fileName: fileName };
}

/** 3PP form = the workbook carries both Offer Analysis and Offer Details tabs. */
function detectFileType_(sheetNames) {
  var names = (sheetNames || []).map(function (name) {
    return String(name || '').toLowerCase().replace(/\s+/g, ' ').trim();
  });

  var hasAnalysis = names.some(function (name) { return name.indexOf('offer analysis') !== -1; });
  var hasDetails = names.some(function (name) { return name.indexOf('offer details') !== -1; });

  return (hasAnalysis && hasDetails) ? FILE_TYPE_3PP : FILE_TYPE_UPDATE;
}

/** Rows logged before the rename still read 'EF'. Show them as 3PP form. */
function normalizeFileType_(value) {
  var text = String(value || '').trim();
  return text.toUpperCase() === 'EF' ? FILE_TYPE_3PP : text;
}

/** One cell, comma separated, however many hotels the form lists. */
function joinHotelNames_(hotelNames) {
  if (!hotelNames || !hotelNames.length) return '';

  var cleaned = [];
  for (var index = 0; index < hotelNames.length && cleaned.length < MAX_HOTEL_NAMES; index += 1) {
    var name = String(hotelNames[index] || '').replace(/\s+/g, ' ').trim();
    if (name) cleaned.push(name);
  }

  return cleaned.join(', ');
}

function stripExtension_(fileName) {
  return String(fileName).replace(/\.[^.]+$/, '');
}

function getLogSheet_() {
  var sheet = SpreadsheetApp.openById(LOG_SPREADSHEET_ID).getSheets()[0];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else if (sheet.getLastColumn() < LOG_HEADERS.length) {
    // a column was added after this sheet was first written
    sheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]).setFontWeight('bold');
  }

  return sheet;
}

function appendLogRow_(row) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    getLogSheet_().appendRow(row);
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------------ *
 * History
 * ------------------------------------------------------------------ */

/**
 * options: { mode: 'days'|'custom', days: 7|14|30, from: 'YYYY-MM-DD',
 *            to: 'YYYY-MM-DD', search: '', type: 'all'|'3PP form'|'update form' }
 */
function getHistory(options) {
  options = options || {};

  var sheet = getLogSheet_();
  var lastRow = sheet.getLastRow();
  var timeZone = Session.getScriptTimeZone();
  if (lastRow < 2) return { rows: [], total: 0, timeZone: timeZone };

  var values = sheet.getRange(2, 1, lastRow - 1, LOG_HEADERS.length).getValues();
  var bounds = resolveRange_(options);
  var tokens = searchTokens_(options.search);
  var wantedType = String(options.type || 'all');
  var rows = [];

  for (var index = values.length - 1; index >= 0; index -= 1) {
    var value = values[index];
    var timestamp = value[0] instanceof Date ? value[0] : new Date(value[0]);
    if (!timestamp || isNaN(timestamp.getTime())) continue;
    if (bounds.from && timestamp < bounds.from) continue;
    if (bounds.to && timestamp > bounds.to) continue;

    var row = {
      timestamp: timestamp.getTime(),
      date: Utilities.formatDate(timestamp, timeZone, 'dd MMM yyyy'),
      time: Utilities.formatDate(timestamp, timeZone, 'HH:mm'),
      email: String(value[1] || ''),
      fileType: normalizeFileType_(value[2]),
      fileName: String(value[3] || ''),
      url: String(value[4] || ''),
      outputs: value[5],
      hotelNames: String(value[6] || '')
    };

    if (wantedType !== 'all' && row.fileType !== wantedType) continue;
    if (!matchesTokens_(tokens, row)) continue;
    rows.push(row);
  }

  return { rows: rows, total: rows.length, timeZone: timeZone };
}

function resolveRange_(options) {
  if (String(options.mode) === 'custom' && (options.from || options.to)) {
    return {
      from: options.from ? startOfDay_(parseDateInput_(options.from)) : null,
      to: options.to ? endOfDay_(parseDateInput_(options.to)) : null
    };
  }

  var days = Number(options.days || 7);
  if (!isFinite(days) || days <= 0) days = 7;

  var from = new Date();
  from.setDate(from.getDate() - (days - 1));
  return { from: startOfDay_(from), to: null };
}

function parseDateInput_(value) {
  var match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value).trim());
  if (match) return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  return new Date(value);
}

function startOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0, 0);
}

function endOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59, 999);
}

/** "Amsterdam & Berlin" has to find "RATE UPDATE_Amsterdam & Berlin ...". */
function normalizeText_(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function searchTokens_(query) {
  var normalized = normalizeText_(query);
  return normalized ? normalized.split(' ') : [];
}

function matchesTokens_(tokens, row) {
  if (!tokens.length) return true;
  var haystack = normalizeText_(
    row.fileName + ' ' + row.email + ' ' + row.fileType + ' ' + row.hotelNames
  );
  for (var index = 0; index < tokens.length; index += 1) {
    if (haystack.indexOf(tokens[index]) === -1) return false;
  }
  return true;
}

/* ------------------------------------------------------------------ *
 * Google Sheet link import (server side, so there is no CORS problem
 * and private sheets the user can open also work)
 * ------------------------------------------------------------------ */

function fetchGoogleSheetAsXlsx(rawUrl) {
  var url = String(rawUrl || '').trim();
  if (!url) throw new Error('Paste a Google Sheet URL first.');
  if (url.indexOf('https://docs.google.com/spreadsheets/') !== 0) {
    throw new Error('Only https://docs.google.com/spreadsheets/ links are accepted.');
  }

  var published = /^https:\/\/docs\.google\.com\/spreadsheets\/d\/e\/([^\/]+)/.exec(url);
  var standard = /^https:\/\/docs\.google\.com\/spreadsheets\/d\/([^\/]+)/.exec(url);

  var exportUrl;
  var fileName;

  if (published) {
    exportUrl = 'https://docs.google.com/spreadsheets/d/e/' + published[1] + '/pub?output=xlsx';
    fileName = 'Google Sheet ' + published[1].slice(0, 8);
  } else if (standard) {
    exportUrl = 'https://docs.google.com/spreadsheets/d/' + standard[1] + '/export?format=xlsx';
    fileName = lookupSheetName_(standard[1]) || ('Google Sheet ' + standard[1].slice(0, 8));
  } else {
    throw new Error('Paste a Google Sheet link from docs.google.com/spreadsheets.');
  }

  var response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
    followRedirects: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Could not export that Sheet (status ' + response.getResponseCode() +
      '). Make sure you have access to it, or share it as public.');
  }

  var bytes = response.getBlob().getBytes();
  if (!(bytes.length > 1 && bytes[0] === 0x50 && bytes[1] === 0x4B)) {
    throw new Error('Google returned a web page instead of an Excel export. Check the link and your access.');
  }

  return { base64: Utilities.base64Encode(bytes), fileName: fileName + '.xlsx' };
}

function lookupSheetName_(fileId) {
  try {
    return Drive.Files.get(fileId, { supportsAllDrives: true }).name;
  } catch (error) {
    return '';
  }
}
