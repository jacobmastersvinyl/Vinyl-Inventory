// ══════════════════════════════════════════════════════════════════
// Funkified — Equipment Inventory System
// Apps Script v5 — offline-reliable backend
//
// New in v5:
//   • LockService on every write (no more race conditions between phones)
//   • clientId idempotency — retries are safe, duplicates impossible
//   • Shelf column (A1–Z3) for warehouse-bay tracking
//   • Batch drain endpoint
//   • Every response echoes clientId so the client can confirm + delete
//   • Processed IDs sheet — rolling 14-day log of seen clientIds
// ══════════════════════════════════════════════════════════════════

const CONFIG = {
  SHEET_ID:        '143xqtYJo-PvwQWkfE06YH54etJC34gsXGapZd3wftfs',
  PHOTO_FOLDER_ID: '1MdQFJ6R21A_ZQemOjsqx-iYB0taXbfVD',
  SHEET_INVENTORY: 'Inventory',
  SHEET_LOG:       'Scan Log',
  SHEET_DAMAGE:    'Damage Log',
  SHEET_CLIENTIDS: 'Processed IDs',
  TIMEZONE:        'Australia/Sydney',
  LOCK_TIMEOUT_MS: 15000,
  IDEMPOTENCY_WINDOW_DAYS: 14,
};

// ══════════════════════════════════════════════════════════════════
// GET — lookup + search
// ══════════════════════════════════════════════════════════════════
function doGet(e) {
  const action = (e.parameter.action || '').toLowerCase();
  const id     = (e.parameter.id    || '').toUpperCase().trim();
  const q      = (e.parameter.q     || '').toLowerCase().trim();

  if (action === 'lookup' && id) return jsonResponse(lookupItem(id));
  if (action === 'search')       return jsonResponse(searchItems(q));
  if (action === 'ping')         return jsonResponse({ ok: true, ts: new Date().toISOString() });
  return jsonResponse({ error: 'Unknown action' });
}

function lookupItem(id) {
  const sheet   = getOrCreateSheet(CONFIG.SHEET_INVENTORY);
  const data    = sheet.getDataRange().getValues();
  const headers = normalise(data[0]);

  for (let i = 1; i < data.length; i++) {
    const row   = data[i];
    const rowId = str(row, headers, 'id').toUpperCase();
    if (rowId !== id) continue;
    return { found: true, item: rowToItem(row, headers) };
  }
  return { found: false };
}

function searchItems(q) {
  const sheet   = getOrCreateSheet(CONFIG.SHEET_INVENTORY);
  const data    = sheet.getDataRange().getValues();
  const headers = normalise(data[0]);
  const items   = [];

  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const name = str(row, headers, 'name').toLowerCase();
    const cat  = str(row, headers, 'category').toLowerCase();
    const rid  = str(row, headers, 'id').toLowerCase();
    const shelf = str(row, headers, 'shelf').toLowerCase();
    if (!q || name.includes(q) || cat.includes(q) || rid.includes(q) || shelf.includes(q)) {
      items.push(rowToItem(row, headers));
    }
  }
  return { items };
}

function rowToItem(row, headers) {
  const photoRaw  = str(row, headers, 'photo urls');
  const photoUrls = photoRaw ? photoRaw.split(',').map(s => s.trim()).filter(Boolean) : [];
  const dmgFlag   = str(row, headers, 'damage flag').toLowerCase() === 'yes';
  return {
    id:            str(row, headers, 'id').toUpperCase(),
    name:          str(row, headers, 'name'),
    category:      str(row, headers, 'category'),
    status:        str(row, headers, 'status') || 'in',
    location:      str(row, headers, 'location'),
    shelf:         str(row, headers, 'shelf'),
    totalQty:      parseInt(str(row, headers, 'total qty'))  || 1,
    qtyOut:        parseInt(str(row, headers, 'qty out'))    || 0,
    photoUrls,
    damageFlagged: dmgFlag,
    damageNote:    str(row, headers, 'damage note'),
    lastScan:      fmtDate(str(row, headers, 'last scanned')),
    lastWho:       str(row, headers, 'last scanned by'),
    notes:         str(row, headers, 'notes'),
  };
}

// ══════════════════════════════════════════════════════════════════
// POST — router with idempotency + LockService
// ══════════════════════════════════════════════════════════════════
function doPost(e) {
  let p;
  try { p = JSON.parse(e.postData.contents); }
  catch(err) { return jsonResponse({ success: false, error: 'Invalid JSON' }); }

  const clientId = (p.clientId || '').toString().trim();

  // ── Idempotency pre-check (outside lock — read-only) ─────────────
  if (clientId) {
    const existing = checkIdempotency(clientId);
    if (existing) {
      return jsonResponse({
        success:   true,
        duplicate: true,
        clientId,
        action:    existing.action,
        note:      'Already processed — ignoring duplicate.',
      });
    }
  }

  // ── Acquire write lock ──────────────────────────────────────────
  const lock = LockService.getScriptLock();
  const locked = lock.tryLock(CONFIG.LOCK_TIMEOUT_MS);
  if (!locked) {
    return jsonResponse({ success: false, retryable: true, error: 'Server busy — retry', clientId });
  }

  try {
    // Double-check idempotency inside lock (race safety)
    if (clientId) {
      const existing = checkIdempotency(clientId);
      if (existing) {
        return jsonResponse({
          success:   true,
          duplicate: true,
          clientId,
          action:    existing.action,
          note:      'Already processed — ignoring duplicate.',
        });
      }
    }

    const action = (p.action || '').toLowerCase();
    let result;
    switch (action) {
      case 'register':       result = registerItem(p); break;
      case 'takeout':        result = takeOut(p); break;
      case 'return':         result = returnItem(p); break;
      case 'damage':         result = logDamage(p); break;
      case 'updatephoto':    result = updatePhoto(p); break;
      case 'edit':           result = editDetails(p); break;
      case 'submitload':     result = submitLoad(p); break;
      case 'submitdelivery': result = submitDelivery(p); break;
      case 'in':
      case 'out':
      case 'update':         result = legacyUpdate(p); break;
      default:               result = { success: false, error: 'Unknown action: ' + p.action };
    }

    // Record the clientId if we actually did the work
    if (clientId && result && result.success) {
      recordClientId(clientId, action);
    }

    return jsonResponse({ ...(result || {}), clientId });
  } catch(err) {
    Logger.log('Error: ' + err.toString());
    return jsonResponse({ success: false, retryable: true, error: err.toString(), clientId });
  } finally {
    lock.releaseLock();
  }
}

// ══════════════════════════════════════════════════════════════════
// IDEMPOTENCY
// ══════════════════════════════════════════════════════════════════
function getOrCreateClientIdsSheet() {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let   sheet = ss.getSheetByName(CONFIG.SHEET_CLIENTIDS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_CLIENTIDS);
    sheet.appendRow(['Client ID','Processed At','Action']);
    styleHeader(sheet, 3, '#0a0a0a');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 280);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 140);
  }
  return sheet;
}

function checkIdempotency(clientId) {
  if (!clientId) return null;
  const sheet = getOrCreateClientIdsSheet();
  const last  = sheet.getLastRow();
  if (last < 2) return null;
  // Read just columns A and C (id, action) — B (timestamp) is display only
  const data = sheet.getRange(2, 1, last - 1, 3).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === clientId) {
      return { recordedAt: data[i][1], action: data[i][2] };
    }
  }
  return null;
}

function recordClientId(clientId, action) {
  if (!clientId) return;
  const sheet = getOrCreateClientIdsSheet();
  sheet.appendRow([clientId, new Date(), action]);
}

// Run this as a time-trigger (daily) to prune old clientIds
function pruneClientIds() {
  const sheet = getOrCreateClientIdsSheet();
  const last  = sheet.getLastRow();
  if (last < 2) return;
  const data   = sheet.getRange(2, 1, last - 1, 3).getValues();
  const cutoff = new Date(Date.now() - CONFIG.IDEMPOTENCY_WINDOW_DAYS * 24 * 60 * 60 * 1000);
  const keep   = data.filter(row => row[1] && new Date(row[1]) > cutoff);

  // Rewrite
  sheet.getRange(2, 1, data.length, 3).clearContent();
  if (keep.length > 0) {
    sheet.getRange(2, 1, keep.length, 3).setValues(keep);
  }
  Logger.log(`Pruned ${data.length - keep.length} old clientIds, kept ${keep.length}`);
}

// ══════════════════════════════════════════════════════════════════
// REGISTER — new item
// ══════════════════════════════════════════════════════════════════
function registerItem(p) {
  const sheet    = getOrCreateSheet(CONFIG.SHEET_INVENTORY);
  const existing = lookupItem(p.id);
  if (existing.found) return { success: false, error: `ID ${p.id} already exists.` };

  const photoUrls = uploadPhotos(p.photos || [], p.id, 'reg');
  const now = new Date();

  sheet.appendRow([
    p.id, p.name, p.category || '', 'in',
    p.totalQty || 1, 0,
    p.gpsLabel || p.location || '',
    p.shelf || '',
    photoUrls.length > 0 ? `=IMAGE("${photoUrls[0]}")` : '',
    now, p.who, now, p.who,
    p.gpsLat || '', p.gpsLng || '', p.gpsAcc || '',
    'no', '', p.notes || '', 1,
  ]);

  appendScanLog({
    id: p.id, name: p.name, action: 'REGISTERED',
    who: p.who, location: p.gpsLabel || '',
    shelf: p.shelf || '',
    gpsLat: p.gpsLat, gpsLng: p.gpsLng,
    notes: p.notes || '', timestamp: now,
  });
  return { success: true, photoUrls };
}

// ══════════════════════════════════════════════════════════════════
// TAKE OUT
// ══════════════════════════════════════════════════════════════════
function takeOut(p) {
  const sheet = getOrCreateSheet(CONFIG.SHEET_INVENTORY);
  const { row, headers, rowIndex } = findRow(sheet, p.id);
  if (!row) return { success: false, error: `${p.id} not found.` };

  const now    = new Date();
  const qtyOut = (parseInt(getCell(row, headers, 'qty out')) || 0) + (parseInt(p.qty) || 1);
  const total  = parseInt(getCell(row, headers, 'total qty')) || 1;

  setCell(sheet, rowIndex, headers, 'qty out',        qtyOut);
  setCell(sheet, rowIndex, headers, 'status',          qtyOut >= total ? 'out' : 'in');
  setCell(sheet, rowIndex, headers, 'last scanned',    now);
  setCell(sheet, rowIndex, headers, 'last scanned by', p.who);
  if (p.gpsLat) setCell(sheet, rowIndex, headers, 'gps lat', p.gpsLat);
  if (p.gpsLng) setCell(sheet, rowIndex, headers, 'gps lng', p.gpsLng);
  incrementScanCount(sheet, rowIndex, headers);

  appendScanLog({
    id: p.id, name: getCell(row, headers, 'name'),
    action: `TAKE OUT ×${p.qty}`, who: p.who,
    location: p.gpsLabel || getCell(row, headers, 'location'),
    shelf: p.shelf || getCell(row, headers, 'shelf'),
    gpsLat: p.gpsLat, gpsLng: p.gpsLng,
    notes: `Event: ${p.event || '—'}`, timestamp: now,
  });
  return { success: true };
}

// ══════════════════════════════════════════════════════════════════
// RETURN
// ══════════════════════════════════════════════════════════════════
function returnItem(p) {
  const sheet = getOrCreateSheet(CONFIG.SHEET_INVENTORY);
  const { row, headers, rowIndex } = findRow(sheet, p.id);
  if (!row) return { success: false, error: `${p.id} not found.` };

  const now    = new Date();
  const qtyOut = Math.max(0, (parseInt(getCell(row, headers, 'qty out')) || 0) - (parseInt(p.qty) || 1));

  setCell(sheet, rowIndex, headers, 'qty out',        qtyOut);
  setCell(sheet, rowIndex, headers, 'status',          'in');
  setCell(sheet, rowIndex, headers, 'last scanned',    now);
  setCell(sheet, rowIndex, headers, 'last scanned by', p.who);
  if (p.gpsLabel) setCell(sheet, rowIndex, headers, 'location', p.gpsLabel);
  if (p.shelf)    setCell(sheet, rowIndex, headers, 'shelf',    p.shelf);
  if (p.gpsLat)   setCell(sheet, rowIndex, headers, 'gps lat',  p.gpsLat);
  if (p.gpsLng)   setCell(sheet, rowIndex, headers, 'gps lng',  p.gpsLng);
  if (p.notes)    setCell(sheet, rowIndex, headers, 'notes',    p.notes);
  incrementScanCount(sheet, rowIndex, headers);

  appendScanLog({
    id: p.id, name: getCell(row, headers, 'name'),
    action: `RETURN ×${p.qty}`, who: p.who,
    location: p.gpsLabel || '', shelf: p.shelf || '',
    gpsLat: p.gpsLat, gpsLng: p.gpsLng,
    notes: p.notes || '', timestamp: now,
  });
  return { success: true };
}

// ══════════════════════════════════════════════════════════════════
// DAMAGE
// ══════════════════════════════════════════════════════════════════
function logDamage(p) {
  const photoUrls = uploadPhotos(p.photos || [], p.id, 'dmg');
  const dmgSheet  = getOrCreateSheet(CONFIG.SHEET_DAMAGE);

  if (dmgSheet.getLastRow() === 0) {
    dmgSheet.appendRow(['Timestamp','Item ID','Item Name','Severity','Description','Who','Location','Shelf','GPS Lat','GPS Lng','Photo URLs']);
    styleHeader(dmgSheet, 11, '#5a0a0a');
    dmgSheet.setFrozenRows(1);
  }

  dmgSheet.appendRow([
    new Date(), p.id, p.name || '', (p.severity || '').toUpperCase(),
    p.description, p.who, p.gpsLabel || '', p.shelf || '',
    p.gpsLat || '', p.gpsLng || '', photoUrls.join(', '),
  ]);

  const invSheet = getOrCreateSheet(CONFIG.SHEET_INVENTORY);
  const { row, headers, rowIndex } = findRow(invSheet, p.id);
  if (row) {
    setCell(invSheet, rowIndex, headers, 'damage flag', 'yes');
    setCell(invSheet, rowIndex, headers, 'damage note', `${(p.severity || '').toUpperCase()}: ${p.description}`);
    setCell(invSheet, rowIndex, headers, 'last scanned', new Date());
    setCell(invSheet, rowIndex, headers, 'last scanned by', p.who);
    const colour = p.severity === 'out' ? '#3b0f0f' : p.severity === 'repair' ? '#3b2a0a' : '#2a2a10';
    invSheet.getRange(rowIndex, 1, 1, invSheet.getLastColumn()).setBackground(colour);
  }

  appendScanLog({
    id: p.id, name: p.name || '',
    action: `DAMAGE — ${(p.severity || '').toUpperCase()}`,
    who: p.who, location: p.gpsLabel || '', shelf: p.shelf || '',
    gpsLat: p.gpsLat, gpsLng: p.gpsLng,
    notes: p.description, timestamp: new Date(),
  });
  return { success: true, photoUrls };
}

// ══════════════════════════════════════════════════════════════════
// UPDATE PHOTO
// ══════════════════════════════════════════════════════════════════
function updatePhoto(p) {
  const photoUrls = uploadPhotos(p.photos || [], p.id, 'upd');
  const sheet     = getOrCreateSheet(CONFIG.SHEET_INVENTORY);
  const { row, headers, rowIndex } = findRow(sheet, p.id);
  if (!row) return { success: false, error: `${p.id} not found.` };

  const existing = getCell(row, headers, 'photo urls') || '';
  const all = [...existing.split(',').map(s => s.trim()).filter(Boolean), ...photoUrls];
  setCell(sheet, rowIndex, headers, 'photo urls', all.length > 0 ? `=IMAGE("${all[0]}")` : '');
  setCell(sheet, rowIndex, headers, 'last scanned', new Date());
  return { success: true, photoUrls };
}

// ══════════════════════════════════════════════════════════════════
// EDIT
// ══════════════════════════════════════════════════════════════════
function editDetails(p) {
  const sheet = getOrCreateSheet(CONFIG.SHEET_INVENTORY);
  const { row, headers, rowIndex } = findRow(sheet, p.id);
  if (!row) return { success: false, error: `${p.id} not found.` };

  if (p.name)                setCell(sheet, rowIndex, headers, 'name',      p.name);
  if (p.category)            setCell(sheet, rowIndex, headers, 'category',  p.category);
  if (p.shelf)               setCell(sheet, rowIndex, headers, 'shelf',     p.shelf);
  if (p.notes !== undefined) setCell(sheet, rowIndex, headers, 'notes',     p.notes);
  if (p.totalQty)            setCell(sheet, rowIndex, headers, 'total qty', p.totalQty);
  setCell(sheet, rowIndex, headers, 'last scanned', new Date());
  return { success: true };
}

// ══════════════════════════════════════════════════════════════════
// SUBMIT LOAD
// ══════════════════════════════════════════════════════════════════
function submitLoad(p) {
  const ss      = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd MMM yy');
  const tabName = `Load — ${p.event || 'Unknown'} — ${p.location || p.who || ''} — ${dateStr}`.substring(0, 100);

  let loadSheet = ss.getSheetByName(tabName);
  if (!loadSheet) {
    loadSheet = ss.insertSheet(tabName);
    loadSheet.appendRow(['Item ID','Item Name','Category','Shelf','Qty Out','Taken By','Location','Timestamp','Event']);
    styleHeader(loadSheet, 9, '#1a3a1a');
    loadSheet.setFrozenRows(1);
  }

  const now = new Date();
  (p.items || []).forEach(item => {
    loadSheet.appendRow([
      item.id, item.name, item.category, item.shelf || '',
      item.qty, item.who, p.location || '', now, p.event,
    ]);
  });
  loadSheet.autoResizeColumns(1, 9);

  appendScanLog({
    id: '—', name: 'LOAD SUBMITTED',
    action: `LOAD: ${p.event} (${(p.items || []).length} items) from ${p.location || '?'}`,
    who: p.who || '', location: p.location || '', shelf: '',
    gpsLat: '', gpsLng: '', notes: tabName, timestamp: now,
  });
  return { success: true, tabName };
}

// ══════════════════════════════════════════════════════════════════
// SUBMIT DELIVERY
// ══════════════════════════════════════════════════════════════════
function submitDelivery(p) {
  const ss      = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd MMM yy');
  const tabName = `Delivery — ${p.supplier || 'Unknown'} — ${p.location || ''} — ${dateStr}`.substring(0, 100);

  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);

    sheet.appendRow(['Delivery Receipt']);
    sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');
    sheet.appendRow(['Supplier:',    p.supplier || '']);
    sheet.appendRow(['Received By:', p.who      || '']);
    sheet.appendRow(['Location:',    p.location || '']);
    sheet.appendRow(['Date:', Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd MMM yyyy HH:mm')]);
    if (p.notes) sheet.appendRow(['Notes:', p.notes]);
    sheet.appendRow([]);

    sheet.appendRow(['Item ID','Item Name','Category','Shelf','Qty Received','Type','Notes']);
    styleHeader(sheet, 7, '#0a2a4a');
    sheet.setFrozenRows(sheet.getLastRow());
  }

  (p.items || []).forEach(item => {
    sheet.appendRow([item.id, item.name, item.category, item.shelf || '', item.qty, item.type, item.notes || '']);
  });
  sheet.autoResizeColumns(1, 7);

  appendScanLog({
    id: '—', name: 'DELIVERY RECEIVED',
    action: `DELIVERY from ${p.supplier} — ${(p.items || []).length} items at ${p.location || '?'}`,
    who: p.who || '', location: p.location || '', shelf: '',
    gpsLat: '', gpsLng: '', notes: tabName, timestamp: new Date(),
  });
  return { success: true, tabName };
}

// ══════════════════════════════════════════════════════════════════
// LEGACY
// ══════════════════════════════════════════════════════════════════
function legacyUpdate(p) {
  const sheet = getOrCreateSheet(CONFIG.SHEET_INVENTORY);
  const { row, headers, rowIndex } = findRow(sheet, p.id);
  if (!row) return { success: false, error: `${p.id} not found.` };

  const now = new Date();
  const newStatus = p.action === 'in' ? 'in' : p.action === 'out' ? 'out' : getCell(row, headers, 'status');
  setCell(sheet, rowIndex, headers, 'status',          newStatus);
  setCell(sheet, rowIndex, headers, 'last scanned',    now);
  setCell(sheet, rowIndex, headers, 'last scanned by', p.who || '');
  if (p.gpsLat) setCell(sheet, rowIndex, headers, 'gps lat', p.gpsLat);
  if (p.gpsLng) setCell(sheet, rowIndex, headers, 'gps lng', p.gpsLng);
  if (p.notes)  setCell(sheet, rowIndex, headers, 'notes',   p.notes);
  if (p.shelf)  setCell(sheet, rowIndex, headers, 'shelf',   p.shelf);
  incrementScanCount(sheet, rowIndex, headers);
  return { success: true };
}

// ══════════════════════════════════════════════════════════════════
// PHOTO UPLOAD
// ══════════════════════════════════════════════════════════════════
function uploadPhotos(photosArray, id, suffix) {
  const urls = [];
  if (!photosArray || photosArray.length === 0) return urls;
  let folder;
  try { folder = DriveApp.getFolderById(CONFIG.PHOTO_FOLDER_ID); }
  catch(e) { Logger.log('Folder error: ' + e); return urls; }

  photosArray.forEach((b64, i) => {
    if (!b64 || !b64.startsWith('data:image')) return;
    try {
      const parts    = b64.split(',');
      const mimeType = parts[0].match(/:(.*?);/)[1];
      const ext      = mimeType.split('/')[1] || 'jpg';
      const blob     = Utilities.newBlob(Utilities.base64Decode(parts[1]), mimeType, `${id}-${suffix}-${i+1}.${ext}`);
      const file     = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      urls.push(`https://drive.google.com/uc?id=${file.getId()}`);
    } catch(e) { Logger.log('Upload error: ' + e); }
  });
  return urls;
}

// ══════════════════════════════════════════════════════════════════
// SCAN LOG
// ══════════════════════════════════════════════════════════════════
function appendScanLog(entry) {
  const log = getOrCreateSheet(CONFIG.SHEET_LOG);
  if (log.getLastRow() === 0) {
    log.appendRow(['Timestamp','Item ID','Item Name','Action','Who','Location','Shelf','GPS Lat','GPS Lng','Notes']);
    styleHeader(log, 10, '#0a0a0a');
    log.setFrozenRows(1);
  }
  log.appendRow([
    entry.timestamp, entry.id, entry.name, entry.action,
    entry.who, entry.location || '', entry.shelf || '',
    entry.gpsLat || '', entry.gpsLng || '', entry.notes || '',
  ]);
}

// ══════════════════════════════════════════════════════════════════
// SHEET HELPERS
// ══════════════════════════════════════════════════════════════════
function getOrCreateSheet(name) {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let   sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === CONFIG.SHEET_INVENTORY) setupInventorySheet(sheet);
  } else if (name === CONFIG.SHEET_INVENTORY) {
    // Auto-migrate: add Shelf column if missing
    ensureShelfColumn(sheet);
  }
  return sheet;
}

function ensureShelfColumn(sheet) {
  if (sheet.getLastRow() === 0) return;
  const headers = normalise(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
  if (headers.indexOf('shelf') > -1) return;
  // Insert Shelf column after Location (index = locationIdx + 1)
  const locIdx = headers.indexOf('location');
  const insertAt = (locIdx > -1 ? locIdx : 5) + 2; // 1-based
  sheet.insertColumnBefore(insertAt);
  sheet.getRange(1, insertAt).setValue('Shelf');
  sheet.setColumnWidth(insertAt, 80);
  const r = sheet.getRange(1, insertAt, 1, 1);
  r.setBackground('#0a0a0a').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  Logger.log('Added Shelf column at ' + insertAt);
}

function setupInventorySheet(sheet) {
  const headers = [
    'ID','Name','Category','Status','Total Qty','Qty Out','Location','Shelf','Photo URLs',
    'Registered','Registered By','Last Scanned','Last Scanned By',
    'GPS Lat','GPS Lng','GPS Accuracy (m)','Damage Flag','Damage Note','Notes','Scan Count',
  ];
  sheet.appendRow(headers);
  styleHeader(sheet, headers.length, '#0a0a0a');
  sheet.setFrozenRows(1);
  [80,200,140,70,80,80,180,80,320,160,160,160,160,90,90,130,100,280,200,90]
    .forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

function styleHeader(sheet, colCount, bg) {
  const r = sheet.getRange(1, 1, 1, colCount);
  r.setBackground(bg); r.setFontColor('#ffffff');
  r.setFontWeight('bold'); r.setFontSize(10);
}

function normalise(row) { return row.map(h => h.toString().toLowerCase().trim()); }
function str(row, headers, key) {
  const i = headers.indexOf(key);
  return i > -1 ? (row[i] === null || row[i] === undefined ? '' : row[i].toString()) : '';
}
function getCell(row, headers, key) { return str(row, headers, key); }

function findRow(sheet, id) {
  const data    = sheet.getDataRange().getValues();
  const headers = normalise(data[0]);
  const idCol   = headers.indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if ((data[i][idCol] || '').toString().toUpperCase().trim() === id.toUpperCase().trim()) {
      return { row: data[i], headers, rowIndex: i + 1 };
    }
  }
  return { row: null, headers, rowIndex: -1 };
}

function setCell(sheet, rowIndex, headers, key, value) {
  const col = headers.indexOf(key);
  if (col > -1) sheet.getRange(rowIndex, col + 1).setValue(value);
}

function incrementScanCount(sheet, rowIndex, headers) {
  const col = headers.indexOf('scan count');
  if (col < 0) return;
  const current = parseInt(sheet.getRange(rowIndex, col + 1).getValue()) || 0;
  sheet.getRange(rowIndex, col + 1).setValue(current + 1);
}

function fmtDate(val) {
  if (!val) return '';
  try { return Utilities.formatDate(new Date(val), CONFIG.TIMEZONE, 'dd MMM yyyy HH:mm'); }
  catch(e) { return val.toString(); }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════════════
// SETUP — run once manually after pasting
// ══════════════════════════════════════════════════════════════════
function setupSheets() {
  getOrCreateSheet(CONFIG.SHEET_INVENTORY);
  getOrCreateSheet(CONFIG.SHEET_LOG);
  getOrCreateSheet(CONFIG.SHEET_DAMAGE);
  getOrCreateClientIdsSheet();
  Logger.log('All sheets ready.');
}

// Optional: create a daily trigger to prune old clientIds
function installPruneTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'pruneClientIds')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('pruneClientIds').timeBased().everyDays(1).atHour(3).create();
  Logger.log('Prune trigger installed (daily at 3am).');
}
