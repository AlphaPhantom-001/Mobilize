/**
 * Mobilize — Google Apps Script Sync & Reconciliation Backend (v2)
 *
 * Supports:
 *  - High-speed chunked batch sync (`POST action=batchAddRows`)
 *  - Legacy single-row sync (`POST action=addRow`) for backward compatibility
 *  - Fast fingerprint reconciliation (`POST action=reconcileCheck`, no full-sheet download)
 *  - Lightweight count & capability discovery (`GET ?action=count` / `?action=reconcileMeta`)
 *  - Incremental delta / full sheet read (`GET ?action=delta&sinceRow=N` / default `GET`)
 *
 * Deployment:
 *  1. Paste this entire file into Code.gs in your Google Apps Script project.
 *  2. Deploy -> Manage deployments -> Edit -> New version -> Execute as: Me, Who has access: Anyone.
 *  3. Save/Update and copy the Web App URL into Mobilize Settings (if the URL changed).
 */

var DEFAULT_SHEET_ID = "186UvLFx_AlyOI7jQXZ0GwyAJQk8k4MFtUmTpkyNMAGg";
var DEFAULT_FOLDER_ID = "1LNK2R4m5Hzl55z1h-WXJJ3W2kx9AWE0w";
var SCRIPT_VERSION = 2;

var DEFAULT_HEADERS = [
  "S/N",
  "NAME",
  "PORTFOLIO/MDA",
  "L.G.A",
  "Code",
  "WARD",
  "Code",
  "POLLING UNIT",
  "Code",
  "CONTACT NO",
  "VIN",
  "IMAGE URL"
];

// ─────────────────────────────────────────────────────────────────
// Deep Normalization Helpers (must stay in sync with Mobilize.html)
// ─────────────────────────────────────────────────────────────────
function normStr(str) {
  return String(str || "").trim().toUpperCase().replace(/[.,'’\s]+/g, ' ').trim();
}

function normVinStr(str) {
  return String(str || "").trim().toUpperCase().replace(/\s+/g, '');
}

function pad3Str(str) {
  return String(str || "").replace(/\D/g, '').padStart(3, '0');
}

function pad2Str(str) {
  return String(str || "").replace(/\D/g, '').padStart(2, '0');
}

function normContactStr(str) {
  var raw = String(str || "").trim();
  if (!raw) return "";
  var digits = raw.replace(/\D/g, '');
  // Normalize 10-digit Nigerian numbers that lost a leading zero in spreadsheets
  if (digits.length === 10 && /^[789]/.test(digits)) {
    return "0" + digits;
  }
  return digits.length >= 10 ? digits : raw;
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────────────
// Sheet & Column Discovery
// ─────────────────────────────────────────────────────────────────
function openTargetSheet_(sheetId) {
  var ss = sheetId ? SpreadsheetApp.openById(sheetId) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Unable to open target Google Sheet.");
  var sheet = ss.getSheets()[0];
  if (!sheet) throw new Error("Target Google Sheet has no worksheets.");
  return sheet;
}

function ensureSheetHeaders_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) {
    var initHeaders = DEFAULT_HEADERS.concat(["PVC"]);
    sheet.getRange(1, 1, 1, initHeaders.length).setValues([initHeaders]);
    sheet.getRange(1, 1, 1, initHeaders.length).setFontWeight("bold");
    return initHeaders;
  }
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
}

function resolveColumnMap_(headers, ensurePvcColumn, sheet) {
  var map = {
    sn: -1,
    name: -1,
    portfolio: -1,
    lga: -1,
    lgaCode: -1,
    ward: -1,
    wardCode: -1,
    pu: -1,
    puCode: -1,
    contact: -1,
    vin: -1,
    imageUrl: -1,
    pvc: -1
  };

  var lastLocationField = "";
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || "").trim().toUpperCase().replace(/\s+/g, " ");
    if (!h) continue;

    if (/^S\/?N$|^S\.?NO$|^SERIAL/.test(h) && map.sn === -1) { map.sn = i; continue; }
    if ((h === "NAME" || /FULL\s*NAME|VOTER\s*NAME/.test(h)) && map.name === -1) { map.name = i; continue; }
    if (/PORTFOLIO|MDA|ROLE|POSITION/.test(h) && map.portfolio === -1) { map.portfolio = i; continue; }

    if (/L\.?G\.?A.*CODE|^LGACODE$/.test(h) && map.lgaCode === -1) { map.lgaCode = i; lastLocationField = "lga"; continue; }
    if (/WARD.*CODE|^RA.*CODE|^WARDCODE$/.test(h) && map.wardCode === -1) { map.wardCode = i; lastLocationField = "ward"; continue; }
    if (/(POLLING\s*UNIT|P\.?U\.?).*CODE|^PUCODE$/.test(h) && map.puCode === -1) { map.puCode = i; lastLocationField = "pu"; continue; }

    if (/L\.?G\.?A|^LGA$|LOCAL\s*GOV/.test(h) && map.lga === -1) { map.lga = i; lastLocationField = "lga"; continue; }
    if (/^WARD$|^WARD\s*NAME$|^RA\/?WARD$/.test(h) && map.ward === -1) { map.ward = i; lastLocationField = "ward"; continue; }
    if (/POLLING\s*UNIT|^P\.?U\.?$/.test(h) && map.pu === -1) { map.pu = i; lastLocationField = "pu"; continue; }

    if (/^CODE(\s*NO\.?)?$/.test(h)) {
      if (lastLocationField === "lga" && map.lgaCode === -1) { map.lgaCode = i; continue; }
      if (lastLocationField === "ward" && map.wardCode === -1) { map.wardCode = i; continue; }
      if (lastLocationField === "pu" && map.puCode === -1) { map.puCode = i; continue; }
      if (map.lgaCode === -1) { map.lgaCode = i; continue; }
      if (map.wardCode === -1) { map.wardCode = i; continue; }
      if (map.puCode === -1) { map.puCode = i; continue; }
    }

    if (/CONTACT|PHONE|MOBILE|GSM/.test(h) && map.contact === -1) { map.contact = i; continue; }
    if (/^VIN$|VOTER.*ID/.test(h) && map.vin === -1) { map.vin = i; continue; }
    if (/^PVC$/.test(h) && map.pvc === -1) { map.pvc = i; continue; }
    if (/IMAGE|PHOTO|URL/.test(h) && map.imageUrl === -1) { map.imageUrl = i; continue; }
  }

  // Fallback to exact 12-column schema if headers are standard or partial
  if (map.sn === -1) map.sn = 0;
  if (map.name === -1) map.name = 1;
  if (map.portfolio === -1) map.portfolio = 2;
  if (map.lga === -1) map.lga = 3;
  if (map.lgaCode === -1) map.lgaCode = 4;
  if (map.ward === -1) map.ward = 5;
  if (map.wardCode === -1) map.wardCode = 6;
  if (map.pu === -1) map.pu = 7;
  if (map.puCode === -1) map.puCode = 8;
  if (map.contact === -1) map.contact = 9;
  if (map.vin === -1) map.vin = 10;
  if (map.imageUrl === -1) map.imageUrl = 11;

  if (ensurePvcColumn && map.pvc === -1 && sheet) {
    var nextCol = Math.max(headers.length, 12) + 1;
    sheet.getRange(1, nextCol).setValue("PVC").setFontWeight("bold");
    headers.push("PVC");
    map.pvc = nextCol - 1;
  }

  map.numCols = Math.max(
    headers.length,
    12,
    map.pvc >= 0 ? map.pvc + 1 : 12
  );
  return map;
}

// ─────────────────────────────────────────────────────────────────
// Sheet Revision Tracking (for cache & fast reconciliation)
// ─────────────────────────────────────────────────────────────────
function getSheetRevision_(sheetId, sheet) {
  var props = PropertiesService.getScriptProperties();
  var key = "rev_" + (sheetId || "default");
  var stored = Number(props.getProperty(key) || "0");
  var rowCount = Math.max(0, sheet.getLastRow() - 1);
  // Combine monotonic mutation counter + live row count so manual edits also bump
  return stored + ":" + rowCount;
}

function bumpSheetRevision_(sheetId, sheet) {
  var props = PropertiesService.getScriptProperties();
  var key = "rev_" + (sheetId || "default");
  var next = Number(props.getProperty(key) || "0") + 1;
  props.setProperty(key, String(next));
  var rowCount = Math.max(0, sheet.getLastRow() - 1);
  return next + ":" + rowCount;
}

// ─────────────────────────────────────────────────────────────────
// Index Building & Fingerprints
// ─────────────────────────────────────────────────────────────────
function registerRowKeys_(rowArr, rowIdx, colMap, index) {
  var vin = normVinStr(rowArr[colMap.vin]);
  var name = normStr(rowArr[colMap.name]);
  var puCode = pad3Str(rowArr[colMap.puCode]);
  var lgaCode = pad2Str(rowArr[colMap.lgaCode]);
  var contact = normContactStr(rowArr[colMap.contact]);

  if (vin && !index.vin.has(vin)) {
    index.vin.set(vin, rowIdx);
  }
  if (name && puCode !== "000") {
    var namePuKey = name + "|" + puCode;
    if (!index.namePu.has(namePuKey)) {
      index.namePu.set(namePuKey, rowIdx);
    }
  }
  if (name && lgaCode !== "00" && contact) {
    var nameLgaContactKey = name + "|" + lgaCode + "|" + contact;
    if (!index.nameLgaContact.has(nameLgaContactKey)) {
      index.nameLgaContact.set(nameLgaContactKey, rowIdx);
    }
  }
}

function buildSheetIndex_(dataValues, colMap) {
  var index = {
    vin: new Map(),
    namePu: new Map(),
    nameLgaContact: new Map()
  };
  for (var i = 0; i < dataValues.length; i++) {
    registerRowKeys_(dataValues[i], i, colMap, index);
  }
  return index;
}

function findMatchingRowIdx_(rec, index) {
  var vin = normVinStr(rec.vin || rec.contact2);
  if (vin && index.vin.has(vin)) {
    return index.vin.get(vin);
  }
  var name = normStr(rec.name);
  var puCode = pad3Str(rec.puCode);
  if (name && puCode !== "000") {
    var k1 = name + "|" + puCode;
    if (index.namePu.has(k1)) {
      return index.namePu.get(k1);
    }
  }
  var lgaCode = pad2Str(rec.lgaCode);
  var contact = normContactStr(rec.contact);
  if (name && lgaCode !== "00" && contact) {
    var k2 = name + "|" + lgaCode + "|" + contact;
    if (index.nameLgaContact.has(k2)) {
      return index.nameLgaContact.get(k2);
    }
  }
  return -1;
}

// ─────────────────────────────────────────────────────────────────
// Drive Image Upload
// ─────────────────────────────────────────────────────────────────
function uploadCardImage_(folder, imageData, fileName) {
  if (!imageData) return "";
  var raw = String(imageData).trim();
  if (!raw) return "";

  var mimeType = "image/jpeg";
  var b64 = raw;
  var m = raw.match(/^data:([^;]+);base64,(.+)$/);
  if (m) {
    mimeType = m[1] || "image/jpeg";
    b64 = m[2];
  }

  var safeName = String(fileName || ("card_" + Date.now() + ".jpg")).replace(/[^\w.\-]+/g, "_");
  var bytes = Utilities.base64Decode(b64);
  var blob = Utilities.newBlob(bytes, mimeType, safeName);
  var file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (shareErr) {
    // Workspace policies may restrict external link sharing; keep file URL
  }
  return file.getUrl();
}

// ─────────────────────────────────────────────────────────────────
// Merge & Write Helpers (No Data Loss on Update, Text-Format Safe)
// ─────────────────────────────────────────────────────────────────
function padRowWidth_(arr, numCols) {
  var out = arr.slice(0, numCols);
  while (out.length < numCols) out.push("");
  return out;
}

function applyTextFormatToColumns_(sheet, startRow, numRows, colMap) {
  if (numRows <= 0) return;
  var textCols = [colMap.lgaCode, colMap.wardCode, colMap.puCode, colMap.contact, colMap.vin];
  if (colMap.pvc >= 0) textCols.push(colMap.pvc);
  for (var i = 0; i < textCols.length; i++) {
    var colIdx = textCols[i];
    if (colIdx >= 0) {
      sheet.getRange(startRow, colIdx + 1, numRows, 1).setNumberFormat("@");
    }
  }
}

function writeContiguousRowUpdates_(sheet, existingData, updatedIndexSet, colMap) {
  if (updatedIndexSet.size === 0) return;
  var sortedIndices = Array.from(updatedIndexSet).sort(function(a, b) { return a - b; });
  var blockStart = sortedIndices[0];
  var blockRows = [padRowWidth_(existingData[blockStart], colMap.numCols)];

  for (var i = 1; i < sortedIndices.length; i++) {
    var idx = sortedIndices[i];
    if (idx === sortedIndices[i - 1] + 1) {
      blockRows.push(padRowWidth_(existingData[idx], colMap.numCols));
    } else {
      applyTextFormatToColumns_(sheet, blockStart + 2, blockRows.length, colMap);
      sheet.getRange(blockStart + 2, 1, blockRows.length, colMap.numCols).setValues(blockRows);
      blockStart = idx;
      blockRows = [padRowWidth_(existingData[idx], colMap.numCols)];
    }
  }
  applyTextFormatToColumns_(sheet, blockStart + 2, blockRows.length, colMap);
  sheet.getRange(blockStart + 2, 1, blockRows.length, colMap.numCols).setValues(blockRows);
}

// ─────────────────────────────────────────────────────────────────
// Core Batch Record Processor (handles both 1-row & N-row batches)
// ─────────────────────────────────────────────────────────────────
function processBatchRecords_(sheetId, folderId, records) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var sheet = openTargetSheet_(sheetId);
    var headers = ensureSheetHeaders_(sheet);
    var hasIncomingPvc = records.some(function(r) { return Boolean(String(r && r.pvc || "").trim()); });
    var colMap = resolveColumnMap_(headers, hasIncomingPvc, sheet);

    var lastRow = sheet.getLastRow();
    var existingData = [];
    if (lastRow >= 2) {
      existingData = sheet.getRange(2, 1, lastRow - 1, colMap.numCols).getValues();
    }

    var index = buildSheetIndex_(existingData, colMap);
    var updatedExistingSet = new Set();
    var newRows = [];
    var allRowsRef = existingData.slice();

    var hasAnyImage = records.some(function(r) { return Boolean(r && r.imageData); });
    var folder = null;
    var folderErr = null;
    if (hasAnyImage) {
      try {
        folder = DriveApp.getFolderById(folderId || DEFAULT_FOLDER_ID);
      } catch (e) {
        folderErr = e;
      }
    }

    // Track highest numeric S/N for clean auto-increment when needed
    var maxSn = existingData.length;
    for (var s = 0; s < existingData.length; s++) {
      var n = Number(existingData[s][colMap.sn]);
      if (isFinite(n) && n > maxSn) maxSn = n;
    }

    var results = [];
    var syncedCount = 0;
    var updatedCount = 0;
    var errorCount = 0;

    for (var i = 0; i < records.length; i++) {
      var rec = records[i] || {};
      var clientId = rec.clientId !== undefined ? rec.clientId : (rec.id !== undefined ? rec.id : String(i));

      try {
        var uploadedUrl = "";
        if (rec.imageData) {
          if (!folder) {
            throw new Error("Drive folder unavailable: " + (folderErr ? folderErr.message : folderId));
          }
          uploadedUrl = uploadCardImage_(folder, rec.imageData, rec.fileName || ("row_" + clientId + ".jpg"));
        } else if (rec.imageUrl) {
          uploadedUrl = String(rec.imageUrl).trim();
        }

        var matchIdx = findMatchingRowIdx_(rec, index);
        var inSn = String(rec.sn || "").trim();
        var inName = String(rec.name || "").trim();
        var inPortfolio = String(rec.portfolio || "").trim();
        var inLga = String(rec.lga || "").trim();
        var inLgaCode = String(rec.lgaCode || "").trim() ? pad2Str(rec.lgaCode) : "";
        var inWard = String(rec.ward || "").trim();
        var inWardCode = String(rec.wardCode || "").trim() ? pad2Str(rec.wardCode) : "";
        var inPu = String(rec.pu || "").trim();
        var inPuCode = String(rec.puCode || "").trim() ? pad3Str(rec.puCode) : "";
        var inContact = String(rec.contact || "").trim();
        var inVin = normVinStr(rec.vin);
        var inPvc = String(rec.pvc || "").trim();

        if (matchIdx >= 0) {
          // UPDATE EXISTING ROW WITHOUT OVERWRITING NON-EMPTY CELLS WITH BLANKS
          var targetRow = padRowWidth_(allRowsRef[matchIdx], colMap.numCols);
          if (inSn && !String(targetRow[colMap.sn] || "").trim()) targetRow[colMap.sn] = inSn;
          if (inName) targetRow[colMap.name] = inName;
          if (inPortfolio) targetRow[colMap.portfolio] = inPortfolio;
          if (inLga) targetRow[colMap.lga] = inLga;
          if (inLgaCode) targetRow[colMap.lgaCode] = inLgaCode;
          if (inWard) targetRow[colMap.ward] = inWard;
          if (inWardCode) targetRow[colMap.wardCode] = inWardCode;
          if (inPu) targetRow[colMap.pu] = inPu;
          if (inPuCode) targetRow[colMap.puCode] = inPuCode;
          if (inContact) targetRow[colMap.contact] = inContact;
          if (inVin) targetRow[colMap.vin] = inVin;
          if (uploadedUrl) targetRow[colMap.imageUrl] = uploadedUrl;
          if (colMap.pvc >= 0 && inPvc) targetRow[colMap.pvc] = inPvc;

          allRowsRef[matchIdx] = targetRow;
          if (matchIdx < existingData.length) {
            existingData[matchIdx] = targetRow;
            updatedExistingSet.add(matchIdx);
          } else {
            newRows[matchIdx - existingData.length] = targetRow;
          }
          registerRowKeys_(targetRow, matchIdx, colMap, index);

          updatedCount++;
          results.push({
            clientId: clientId,
            result: "updated",
            sn: targetRow[colMap.sn],
            rowNumber: matchIdx + 2,
            imageUrl: String(targetRow[colMap.imageUrl] || uploadedUrl || "")
          });
        } else {
          // APPEND NEW ROW
          maxSn++;
          var finalSn = inSn || String(maxSn);
          var newRow = new Array(colMap.numCols).fill("");
          newRow[colMap.sn] = finalSn;
          newRow[colMap.name] = inName;
          newRow[colMap.portfolio] = inPortfolio;
          newRow[colMap.lga] = inLga;
          newRow[colMap.lgaCode] = inLgaCode;
          newRow[colMap.ward] = inWard;
          newRow[colMap.wardCode] = inWardCode;
          newRow[colMap.pu] = inPu;
          newRow[colMap.puCode] = inPuCode;
          newRow[colMap.contact] = inContact;
          newRow[colMap.vin] = inVin;
          newRow[colMap.imageUrl] = uploadedUrl;
          if (colMap.pvc >= 0) newRow[colMap.pvc] = inPvc;

          var combinedIdx = allRowsRef.length;
          allRowsRef.push(newRow);
          newRows.push(newRow);
          registerRowKeys_(newRow, combinedIdx, colMap, index);

          syncedCount++;
          results.push({
            clientId: clientId,
            result: "success",
            sn: finalSn,
            rowNumber: combinedIdx + 2,
            imageUrl: uploadedUrl
          });
        }
      } catch (recErr) {
        errorCount++;
        results.push({
          clientId: clientId,
          result: "error",
          message: recErr && recErr.message ? recErr.message : String(recErr)
        });
      }
    }

    // Apply all sheet updates in grouped ranges
    writeContiguousRowUpdates_(sheet, existingData, updatedExistingSet, colMap);

    // Append all new rows in ONE single setValues() call
    if (newRows.length > 0) {
      var startAppendRow = Math.max(2, sheet.getLastRow() + 1);
      var paddedNewRows = newRows.map(function(r) { return padRowWidth_(r, colMap.numCols); });
      applyTextFormatToColumns_(sheet, startAppendRow, paddedNewRows.length, colMap);
      sheet.getRange(startAppendRow, 1, paddedNewRows.length, colMap.numCols).setValues(paddedNewRows);
    }

    SpreadsheetApp.flush();

    var sheetRevision = bumpSheetRevision_(sheetId, sheet);
    var sheetTotal = Math.max(0, sheet.getLastRow() - 1);

    return {
      result: "ok",
      version: SCRIPT_VERSION,
      processed: records.length,
      syncedCount: syncedCount,
      updatedCount: updatedCount,
      errorCount: errorCount,
      sheetTotal: sheetTotal,
      sheetRevision: sheetRevision,
      results: results
    };
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────────
// Fast Fingerprint Reconciliation (avoids sending full sheet to client)
// ─────────────────────────────────────────────────────────────────
function handleReconcileCheck_(sheetId, fingerprints) {
  var sheet = openTargetSheet_(sheetId);
  var lastRow = sheet.getLastRow();
  var headers = lastRow >= 1 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
  var colMap = resolveColumnMap_(headers, false, null);
  var sheetTotal = Math.max(0, lastRow - 1);
  var sheetRevision = getSheetRevision_(sheetId, sheet);

  var existingData = [];
  if (lastRow >= 2) {
    existingData = sheet.getRange(2, 1, lastRow - 1, colMap.numCols).getValues();
  }
  var index = buildSheetIndex_(existingData, colMap);

  var missingIds = [];
  var matchedMeta = {};

  for (var i = 0; i < fingerprints.length; i++) {
    var fp = fingerprints[i] || {};
    var idx = findMatchingRowIdx_(fp, index);
    if (idx >= 0) {
      var sheetRow = existingData[idx];
      matchedMeta[String(fp.id)] = {
        rowNumber: idx + 2,
        imageUrl: String(sheetRow[colMap.imageUrl] || "")
      };
    } else {
      missingIds.push(fp.id);
    }
  }

  return {
    result: "ok",
    version: SCRIPT_VERSION,
    sheetTotal: sheetTotal,
    sheetRevision: sheetRevision,
    checkedCount: fingerprints.length,
    matchedCount: fingerprints.length - missingIds.length,
    missingIds: missingIds,
    matchedMeta: matchedMeta
  };
}

// ─────────────────────────────────────────────────────────────────
// Sheet Row Serializer (for delta / legacy GET reconciliation)
// ─────────────────────────────────────────────────────────────────
function readSheetRows_(sheetId, sinceRow) {
  var sheet = openTargetSheet_(sheetId);
  var lastRow = sheet.getLastRow();
  var total = Math.max(0, lastRow - 1);
  var sheetRevision = getSheetRevision_(sheetId, sheet);

  if (lastRow < 2) {
    return { rows: [], total: 0, sheetRevision: sheetRevision, sinceRow: 0, nextRow: 0 };
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colMap = resolveColumnMap_(headers, false, null);
  var startOffset = Math.max(0, Number(sinceRow || 0));
  if (startOffset >= total) {
    return { rows: [], total: total, sheetRevision: sheetRevision, sinceRow: startOffset, nextRow: total };
  }

  var numToRead = total - startOffset;
  var values = sheet.getRange(2 + startOffset, 1, numToRead, colMap.numCols).getValues();
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    rows.push({
      rowNumber: startOffset + i + 2,
      sn: r[colMap.sn],
      name: String(r[colMap.name] || ""),
      portfolio: String(r[colMap.portfolio] || ""),
      lga: String(r[colMap.lga] || ""),
      lgaCode: String(r[colMap.lgaCode] || ""),
      ward: String(r[colMap.ward] || ""),
      wardCode: String(r[colMap.wardCode] || ""),
      pu: String(r[colMap.pu] || ""),
      puCode: String(r[colMap.puCode] || ""),
      contact: String(r[colMap.contact] || ""),
      vin: String(r[colMap.vin] || ""),
      imageUrl: String(r[colMap.imageUrl] || ""),
      pvc: colMap.pvc >= 0 ? String(r[colMap.pvc] || "") : ""
    });
  }

  return {
    rows: rows,
    total: total,
    sheetRevision: sheetRevision,
    sinceRow: startOffset,
    nextRow: startOffset + rows.length
  };
}

// ─────────────────────────────────────────────────────────────────
// Web App Entrypoints: doGet & doPost
// ─────────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var action = String(params.action || "").trim();
    var sheetId = String(params.sheetId || DEFAULT_SHEET_ID).trim();

    if (action === "count" || action === "reconcileMeta" || action === "capabilities") {
      var sheet = openTargetSheet_(sheetId);
      var total = Math.max(0, sheet.getLastRow() - 1);
      var rev = getSheetRevision_(sheetId, sheet);
      return jsonOutput_({
        ok: true,
        version: SCRIPT_VERSION,
        total: total,
        rowCount: total,
        sheetRevision: rev,
        capabilities: {
          batchAddRows: true,
          reconcileCheck: true,
          reconcileMeta: true,
          delta: true
        }
      });
    }

    if (action === "delta") {
      var sinceRow = Number(params.sinceRow || 0);
      var deltaRes = readSheetRows_(sheetId, sinceRow);
      return jsonOutput_(deltaRes);
    }

    // Legacy full-sheet read (kept compatible)
    var fullRes = readSheetRows_(sheetId, 0);
    return jsonOutput_(fullRes);
  } catch (err) {
    return jsonOutput_({
      result: "error",
      ok: false,
      message: err && err.message ? err.message : String(err)
    });
  }
}

function doPost(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var bodyJson = null;

    if ((!params.action || params.action === "") && e && e.postData && e.postData.contents) {
      var cType = String(e.postData.type || "");
      if (cType.indexOf("json") !== -1 || e.postData.contents.trim().charAt(0) === "{") {
        try { bodyJson = JSON.parse(e.postData.contents); } catch (_) {}
      }
    }

    var action = String((bodyJson && bodyJson.action) || params.action || "addRow").trim();
    var sheetId = String((bodyJson && bodyJson.sheetId) || params.sheetId || DEFAULT_SHEET_ID).trim();
    var folderId = String((bodyJson && bodyJson.folderId) || params.folderId || DEFAULT_FOLDER_ID).trim();

    if (action === "batchAddRows") {
      var rawRowsJson = (bodyJson && bodyJson.rows) ? bodyJson.rows : JSON.parse(params.rowsJson || "[]");
      if (!Array.isArray(rawRowsJson)) {
        return jsonOutput_({ result: "error", message: "Invalid rowsJson payload; expected array." });
      }
      var batchRes = processBatchRecords_(sheetId, folderId, rawRowsJson);
      return jsonOutput_(batchRes);
    }

    if (action === "reconcileCheck") {
      var fps = (bodyJson && bodyJson.fingerprints) ? bodyJson.fingerprints : JSON.parse(params.fingerprintsJson || "[]");
      if (!Array.isArray(fps)) {
        return jsonOutput_({ result: "error", message: "Invalid fingerprintsJson payload; expected array." });
      }
      var recRes = handleReconcileCheck_(sheetId, fps);
      return jsonOutput_(recRes);
    }

    // Legacy single-row action=addRow (delegates cleanly to shared batch engine)
    var single = bodyJson || {
      clientId: params.clientId || params.id || params.sn || "single",
      sn: params.sn || "",
      name: params.name || "",
      portfolio: params.portfolio || "",
      lga: params.lga || "",
      lgaCode: params.lgaCode || "",
      ward: params.ward || "",
      wardCode: params.wardCode || "",
      pu: params.pu || "",
      puCode: params.puCode || "",
      contact: params.contact || "",
      vin: params.vin || "",
      pvc: params.pvc || "",
      imageData: params.imageData || "",
      fileName: params.fileName || ""
    };

    var oneRes = processBatchRecords_(sheetId, folderId, [single]);
    var first = (oneRes.results && oneRes.results[0]) ? oneRes.results[0] : { result: "error", message: "No result" };
    return jsonOutput_({
      result: first.result,
      sn: first.sn || "",
      imageUrl: first.imageUrl || "",
      message: first.message || "",
      sheetTotal: oneRes.sheetTotal,
      sheetRevision: oneRes.sheetRevision
    });
  } catch (err) {
    return jsonOutput_({
      result: "error",
      message: err && err.message ? err.message : String(err)
    });
  }
}
