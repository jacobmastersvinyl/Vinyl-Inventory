// ═══════════════════════════════════════════════════════════════
// FUNKIFIED — Storage Apps Script v9
// Changes vs v7:
//   • Shelf is no longer a primary key — it's a *location tag*. A single
//     shelf code (e.g. D1) can hold many items, each with its own row.
//     Uniqueness is now (Shelf, Name): one row per item-name on a shelf.
//   • lookupByShelf returns a LIST of items on a shelf instead of a
//     single row. The client decides which one to open (or to add a new
//     item to that shelf).
//   • All per-item actions (takeOut, return, damage, photo, edit, load)
//     resolve the target row via rowIdx → (shelf,name) → shelf-if-unique
//     in that order. Ambiguous shelf-only calls return {ambiguous:true,
//     candidates:[…]} so the caller can disambiguate.
//   • actRegister no longer refuses non-matching shelves — it appends a
//     new row when (shelf, name) doesn't already exist. Top-up still
//     happens for an exact (shelf, name) match.
//   • allowOverwrite payload field is now ignored (kept for backwards-
//     compat — old client builds in flight).
// Changes vs v6 (carried over):
//   • doGet adds `action=search` for fuzzy inventory lookup
//   • searchInventoryItems() returns rowIdx + storedName so the
//     client can route by a stable key and detect shelf reassignment
//   • JobPacking router is dispatched inside doPost if present
//
// Schema (must match your workbook):
//   Inventory:   Shelf | Name | Brand | Type | Status | Total Qty | Qty Out |
//                Photos | Registered | Registered By | Last Scanned |
//                Last Scanned By | GPS Lat | GPS Lng | GPS Accuracy (m) | Notes
//   Damage Log:  Timestamp | Item ID | Item Name | Severity | Description |
//                Who | Location | GPS Lat | GPS Lng | Photo URLs
//   Scan Log:    Timestamp | Shelf | Name | Action | Who | Location |
//                GPS Lat | GPS Lng | Notes
//   Processed IDs (auto-created, hidden): ClientId | Timestamp | Action
// Uniqueness: (Shelf, Name). Photo filenames in Drive use the item Name.
// ═══════════════════════════════════════════════════════════════

const CONFIG = {
  SHEET_ID: '143xqtYJo-PvwQWkfE06YH54etJC34gsXGapZd3wftfs',
  PHOTO_FOLDER_ID: '1MdQFJ6R21A_ZQemOjsqx-iYB0taXbfVD',
  SHEET_INVENTORY: 'Inventory',
  SHEET_DAMAGE: 'Damage Log',
  SHEET_SCAN: 'Scan Log',
  SHEET_CLIENTIDS: 'Processed IDs',
  LOCK_TIMEOUT_MS: 15000,
  IDEMPOTENCY_WINDOW_DAYS: 14
};

const INV_COLS = {
  Shelf: 1, Name: 2, Brand: 3, Type: 4, Status: 5,
  TotalQty: 6, QtyOut: 7, Photos: 8,
  Registered: 9, RegisteredBy: 10,
  LastScanned: 11, LastScannedBy: 12,
  GpsLat: 13, GpsLng: 14, GpsAcc: 15,
  Notes: 16,
  City: 17
};
const INV_HEADER = ['Shelf','Name','Brand','Type','Status','Total Qty','Qty Out','Photos','Registered','Registered By','Last Scanned','Last Scanned By','GPS Lat','GPS Lng','GPS Accuracy (m)','Notes','City'];
const DAMAGE_COL_PHOTOS = 10;
const DAMAGE_HEADER = ['Timestamp','Item ID','Item Name','Severity','Description','Who','Location','GPS Lat','GPS Lng','Photo URLs'];
const SCAN_HEADER = ['Timestamp','Shelf','Name','Action','Who','Location','GPS Lat','GPS Lng','Notes'];
const CLIENTIDS_HEADER = ['ClientId','Timestamp','Action'];

// ─────────── Entry points ───────────
function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || '').toLowerCase();
  const params = (e && e.parameter) || {};
  try {
    // ClientPortal GET (e.g. clientverify). Shim returns a TextOutput or null.
    if (typeof clientPortalGet === 'function') {
      const r = clientPortalGet(action, params);
      if (r) return r;
    }
    // JobPacking GET (none today, kept for future use).
    if (typeof jobPackGet === 'function') {
      const r = jobPackGet(action, params);
      if (r) return r;
    }
    if (action === 'lookup') {
      const shelf = String((params.shelf || params.id || '')).trim().toUpperCase();
      const expectedName = String(params.expectedName || '').trim();
      return jsonOut(lookupByShelf(shelf, { expectedName }));
    }
    if (action === 'search') {
      const q = String(params.q || '').trim();
      return jsonOut(searchInventoryItems(q));
    }
    return jsonOut({ success: true, status: 'ok' });
  } catch (err) {
    return jsonOut({ success: false, error: String(err) });
  }
}

function doPost(e) {
  const raw = e && e.postData && e.postData.contents;
  let body;
  try { body = JSON.parse(raw); } catch (_) { return jsonOut({ success: false, error: 'Bad JSON' }); }
  const clientId = String(body.clientId || '').trim();
  if (!clientId) return jsonOut({ success: false, error: 'Missing clientId' });

  const actionLc = String(body.action || '').toLowerCase().trim();

  // ClientPortal POST (clientcatalogue, clientorder, …). Runs OUTSIDE the
  // script lock — catalogue reads are idempotent and order writes manage
  // their own sheet. Shim returns a TextOutput or null.
  try {
    if (typeof clientPortalPost === 'function') {
      const r = clientPortalPost(actionLc, body);
      if (r) return r;
    }
  } catch (err) {
    return jsonOut({ success: false, error: String(err && err.message || err), clientId });
  }

  const lock = LockService.getScriptLock();
  try { lock.waitLock(CONFIG.LOCK_TIMEOUT_MS); }
  catch (_) { return jsonOut({ success: false, error: 'Server busy; please retry', retryable: true, clientId }); }

  try {
    if (checkIdempotency(clientId)) {
      return jsonOut({ success: true, duplicate: true, clientId });
    }

    // JobPacking POST (jobspending, jobpack, jobcomplete). Shim returns
    // a TextOutput or null; if handled, record idempotency and return.
    try {
      if (typeof jobPackPost === 'function') {
        const out = jobPackPost(actionLc, body);
        if (out) {
          try {
            const parsed = JSON.parse(out.getContent());
            if (parsed && parsed.success !== false) recordClientId(clientId, body.action);
          } catch (_) {}
          return out;
        }
      }
    } catch (_) {}

    const action = String(body.action || '').trim();
    let result;
    switch (action) {
      case 'register':      result = actRegister(body); break;
      case 'registerNoTag': result = actRegisterNoTag(body); break;
      case 'takeOut':       result = actTakeOut(body); break;
      case 'return':        result = actReturn(body); break;
      case 'damage':        result = actDamage(body); break;
      case 'updatePhoto':   result = actUpdatePhoto(body); break;
      case 'edit':          result = actEdit(body); break;
      case 'load':          result = actLoad(body); break;
      case 'delivery':      result = actDelivery(body); break;
      default: return jsonOut({ success: false, error: 'Unknown action: ' + action, clientId });
    }
    if (result && result.success !== false) recordClientId(clientId, action);
    result.clientId = clientId;
    return jsonOut(result);
  } catch (err) {
    return jsonOut({ success: false, error: String(err && err.message || err), clientId });
  } finally {
    lock.releaseLock();
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────── Sheet helpers ───────────
function ss() { return SpreadsheetApp.openById(CONFIG.SHEET_ID); }
function getSheet(name) {
  const s = ss().getSheetByName(name);
  if (!s) throw new Error('Sheet not found: ' + name);
  return s;
}
function ensureHeaders(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return;
  }
  const maxCols = Math.max(sheet.getLastColumn(), headers.length);
  const row = sheet.getRange(1, 1, 1, maxCols).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (String(row[i] || '').trim() !== headers[i]) {
      sheet.getRange(1, i + 1).setValue(headers[i]).setFontWeight('bold');
    }
  }
  sheet.setFrozenRows(1);
}
function normShelf(v) { return String(v || '').trim().toUpperCase(); }
function normName(v)  { return String(v || '').trim().toLowerCase(); }
function nowDate() { return new Date(); }
function locLabel(p) {
  if (p.locationLabel) return p.locationLabel;
  if (p.gpsLat && p.gpsLng) return p.gpsLat + ', ' + p.gpsLng;
  return '';
}

// ─────────── Idempotency ───────────
function getOrCreateClientIdsSheet() {
  const sp = ss();
  let s = sp.getSheetByName(CONFIG.SHEET_CLIENTIDS);
  if (!s) {
    s = sp.insertSheet(CONFIG.SHEET_CLIENTIDS);
    s.getRange(1, 1, 1, CLIENTIDS_HEADER.length).setValues([CLIENTIDS_HEADER]).setFontWeight('bold');
    try { s.hideSheet(); } catch (_) {}
  }
  return s;
}
function checkIdempotency(clientId) {
  const s = getOrCreateClientIdsSheet();
  const last = s.getLastRow();
  if (last < 2) return false;
  const ids = s.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === clientId) return true;
  return false;
}
function recordClientId(clientId, action) {
  const s = getOrCreateClientIdsSheet();
  s.appendRow([clientId, new Date(), action]);
}
function pruneClientIds() {
  const s = getOrCreateClientIdsSheet();
  const last = s.getLastRow();
  if (last < 2) return;
  const vals = s.getRange(2, 1, last - 1, CLIENTIDS_HEADER.length).getValues();
  const cutoff = Date.now() - CONFIG.IDEMPOTENCY_WINDOW_DAYS * 86400000;
  const keep = vals.filter(function (r) {
    return r[1] instanceof Date ? r[1].getTime() >= cutoff : true;
  });
  s.getRange(2, 1, last - 1, CLIENTIDS_HEADER.length).clearContent();
  if (keep.length) s.getRange(2, 1, keep.length, CLIENTIDS_HEADER.length).setValues(keep);
}

// ─────────── Lookup / Row helpers ───────────
// Shelf is a *location tag*, not a key. Multiple items may live on one
// shelf — each is a separate row. The (Shelf, Name) pair is what's
// unique. rowIdx is the stable handle the client should pass back when
// it has one, since it survives shelf moves and renames.
function findRowsByShelf(invSheet, shelf) {
  const want = normShelf(shelf);
  if (!want) return [];
  const last = invSheet.getLastRow();
  if (last < 2) return [];
  const shelves = invSheet.getRange(2, INV_COLS.Shelf, last - 1, 1).getValues();
  const out = [];
  for (let i = 0; i < shelves.length; i++) {
    if (normShelf(shelves[i][0]) === want) out.push(i + 2);
  }
  return out;
}
function findRowByShelfAndName(invSheet, shelf, name) {
  const wantShelf = normShelf(shelf);
  const wantName = normName(name);
  if (!wantShelf || !wantName) return -1;
  const last = invSheet.getLastRow();
  if (last < 2) return -1;
  const vals = invSheet.getRange(2, 1, last - 1, 2).getValues(); // Shelf, Name
  for (let i = 0; i < vals.length; i++) {
    if (normShelf(vals[i][0]) === wantShelf && normName(vals[i][1]) === wantName) {
      return i + 2;
    }
  }
  return -1;
}
// Back-compat: the old single-result lookup still works but now returns
// the FIRST row on the shelf. New code should prefer findRowsByShelf or
// findRowByShelfAndName. Kept so any half-deployed client still functions.
function findRowByShelf(invSheet, shelf) {
  const list = findRowsByShelf(invSheet, shelf);
  return list.length ? list[0] : -1;
}
function readRow(invSheet, rowIdx) {
  const v = invSheet.getRange(rowIdx, 1, 1, INV_HEADER.length).getValues()[0];
  return {
    shelf: v[0], name: v[1], brand: v[2], type: v[3], status: v[4],
    totalQty: v[5], qtyOut: v[6], photos: v[7],
    registered: v[8], registeredBy: v[9],
    lastScanned: v[10], lastScannedBy: v[11],
    gpsLat: v[12], gpsLng: v[13], gpsAcc: v[14],
    notes: v[15],
    rowIdx: rowIdx
  };
}
// Resolve a target row from a request payload. Tries rowIdx, then
// (shelf,name), then bare shelf. If shelf-only resolves to multiple
// items, returns ambiguous so the caller can pick.
function resolveRow(invSheet, p) {
  const last = invSheet.getLastRow();
  // 1) explicit rowIdx
  const rIdx = Number(p.rowIdx) || 0;
  if (rIdx >= 2 && rIdx <= last) return { row: rIdx };
  // 2) (shelf, name)
  const shelf = normShelf(p.shelf || p.id);
  const name  = String(p.name || '').trim();
  if (shelf && name) {
    const r = findRowByShelfAndName(invSheet, shelf, name);
    if (r >= 0) return { row: r };
  }
  // 3) bare shelf — only if a single row sits on it
  if (shelf) {
    const list = findRowsByShelf(invSheet, shelf);
    if (list.length === 1) return { row: list[0] };
    if (list.length > 1) {
      const candidates = list.map(function (rn) {
        const it = readRow(invSheet, rn);
        return { rowIdx: rn, name: it.name, brand: it.brand, type: it.type, status: it.status };
      });
      return { row: -1, ambiguous: true, candidates: candidates, shelf: shelf };
    }
  }
  return { row: -1 };
}
// New shape: returns { exists, shelf, items: [...] }
//   • items is an array (possibly empty) of all rows on this shelf
//   • when expectedName is supplied, items is filtered to that name
//     and `mismatch:true` is set if no such row exists on the shelf
function lookupByShelf(shelf, opts) {
  const s = getSheet(CONFIG.SHEET_INVENTORY);
  ensureHeaders(s, INV_HEADER);
  const rows = findRowsByShelf(s, shelf);
  if (!rows.length) return { exists: false, shelf: shelf, items: [] };
  const items = rows.map(function (r) { return readRow(s, r); });
  const expectedName = String((opts && opts.expectedName) || '').trim();
  if (expectedName) {
    const matched = items.filter(function (it) {
      return normName(it.name) === normName(expectedName);
    });
    if (!matched.length) {
      // Shelf has things on it, but not the item the caller was expecting.
      // (Most often: search snapshot was stale.) Surface so the client
      // can show the user what's actually here.
      return {
        exists: true,
        mismatch: true,
        shelf: shelf,
        expectedName: expectedName,
        items: items
      };
    }
    return { exists: true, shelf: shelf, items: matched };
  }
  return { exists: true, shelf: shelf, items: items };
}

// ─────────── Fuzzy search ───────────
// Returns all rows whose shelf / name / brand / type contain any search
// token (case-insensitive, substring). Exact-shelf matches surface first.
// The client receives rowIdx + storedName so it can detect drift when the
// user later clicks an item that may have been overwritten or edited.
function searchInventoryItems(q) {
  const term = String(q || '').trim();
  if (!term) return { success: true, items: [] };
  const s = getSheet(CONFIG.SHEET_INVENTORY);
  ensureHeaders(s, INV_HEADER);
  const last = s.getLastRow();
  if (last < 2) return { success: true, items: [] };
  const rows = s.getRange(2, 1, last - 1, INV_HEADER.length).getValues();
  // Photo formulas + notes — used to attach a thumb URL to each result.
  const photoFx = s.getRange(2, INV_COLS.Photos, last - 1, 1).getFormulas();
  const photoNts = s.getRange(2, INV_COLS.Photos, last - 1, 1).getNotes();
  const wantShelf = normShelf(term);
  const tokens = term.toLowerCase().split(/\s+/).filter(Boolean);
  const hay = (r) => [r[0], r[1], r[2], r[3]].map(v => String(v || '').toLowerCase()).join(' \u0001 ');
  const scored = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const h = hay(r);
    const shelf = normShelf(r[0]);
    let score = 0;
    if (shelf && shelf === wantShelf) score += 100;
    else if (shelf && shelf.indexOf(wantShelf) === 0 && wantShelf) score += 50;
    let hits = 0;
    for (const t of tokens) {
      if (!t) continue;
      if (h.indexOf(t) !== -1) hits++;
    }
    if (hits === 0 && score === 0) continue;
    score += hits * 10;
    scored.push({ i, score, row: r });
  }
  scored.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    const an = String(a.row[1] || '').toLowerCase();
    const bn = String(b.row[1] || '').toLowerCase();
    return an < bn ? -1 : (an > bn ? 1 : 0);
  });
  const items = scored.slice(0, 80).map(x => {
    const fx = String(photoFx[x.i][0] || '');
    const note = String(photoNts[x.i][0] || '');
    let m = /id=([A-Za-z0-9_-]+)/.exec(fx);
    if (!m && note) m = /id=([A-Za-z0-9_-]+)/.exec(note.split('\n')[0] || '');
    const thumbUrl = m ? ('https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w600') : '';
    return {
      rowIdx: x.i + 2,
      shelf: x.row[0],
      name: x.row[1],
      brand: x.row[2],
      type: x.row[3],
      status: x.row[4],
      totalQty: x.row[5],
      qtyOut: x.row[6],
      notes: x.row[15],
      thumbUrl: thumbUrl
    };
  });
  return { success: true, items };
}

// ─────────── Photo handling ───────────
function slug(s) {
  return String(s || '')
    .replace(/[\/\\:\*\?"<>\|]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}
function savePhoto(dataUrl, name) {
  if (!dataUrl) return null;
  const folder = DriveApp.getFolderById(CONFIG.PHOTO_FOLDER_ID);
  let contentType = 'image/jpeg';
  let b64 = dataUrl;
  const m = /^data:([^;]+);base64,(.+)$/.exec(dataUrl);
  if (m) { contentType = m[1]; b64 = m[2]; }
  const bytes = Utilities.base64Decode(b64);
  const ext = contentType.indexOf('png') !== -1 ? 'png' : 'jpg';
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'UTC', "yyyy-MM-dd'T'HH-mm-ss");
  const base = slug(name) || 'unnamed';
  const fname = base + ' — ' + ts + '.' + ext;
  const blob = Utilities.newBlob(bytes, contentType, fname);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const id = file.getId();
  return {
    id: id,
    url: 'https://drive.google.com/uc?id=' + id,
    thumb: 'https://drive.google.com/thumbnail?id=' + id + '&sz=w600'
  };
}

function renderPhotoCell(sheet, row, col, photo) {
  if (!photo) return;
  const cell = sheet.getRange(row, col);
  cell.setFormula('=IMAGE("' + photo.thumb + '")');
  const existing = cell.getNote() || '';
  cell.setNote(existing ? (existing + '\n' + photo.url) : photo.url);
  try { sheet.setRowHeight(row, 120); } catch (_) {}
}

// ─────────── Scan log ───────────
function appendScan(p, shelf, name, action, notes) {
  const s = getSheet(CONFIG.SHEET_SCAN);
  ensureHeaders(s, SCAN_HEADER);
  s.appendRow([
    nowDate(), shelf || '', name || '', action || '',
    p.team || '', locLabel(p),
    p.gpsLat || '', p.gpsLng || '', notes || ''
  ]);
}

// ─────────── Actions ───────────
// actRegister:
//   • A shelf can hold many items now. Uniqueness is (Shelf, Name).
//   • If (shelf, name) already exists → top-up that row (re-photo, edit
//     brand/type/qty/notes).
//   • Otherwise → append a new row, even if the shelf already has other
//     items on it.
//   • allowOverwrite is accepted but ignored (no overwrite path remains —
//     a different name on the same shelf is just a different item).
function actRegister(p) {
  const inv = getSheet(CONFIG.SHEET_INVENTORY); ensureHeaders(inv, INV_HEADER);
  const shelf = normShelf(p.shelf);
  if (!shelf) return { success: false, error: 'Shelf is required' };
  if (!p.name) return { success: false, error: 'Name is required' };

  const now = nowDate();
  let row = findRowByShelfAndName(inv, shelf, p.name);

  if (row < 0) {
    // No existing item with this (shelf, name) — append a new row.
    const photo = p.photo ? savePhoto(p.photo, p.name) : null;
    inv.appendRow([
      shelf, p.name || '', p.brand || '', p.type || '', 'in',
      Number(p.totalQty) || 0, 0, '',
      now, p.team || '', now, p.team || '',
      p.gpsLat || '', p.gpsLng || '', p.gpsAccuracy || '',
      p.notes || ''
    ]);
    row = inv.getLastRow();
    if (photo) renderPhotoCell(inv, row, INV_COLS.Photos, photo);
    appendScan(p, shelf, p.name, 'REGISTERED', p.notes || '');
    return { success: true, shelf: shelf, name: p.name, rowIdx: row, created: true };
  }

  // Exact (shelf, name) match — top-up.
  const cur = readRow(inv, row);
  const photo = p.photo ? savePhoto(p.photo, p.name) : null;
  if (p.brand) inv.getRange(row, INV_COLS.Brand).setValue(p.brand);
  if (p.type)  inv.getRange(row, INV_COLS.Type).setValue(p.type);
  if (p.totalQty !== undefined && p.totalQty !== '') {
    inv.getRange(row, INV_COLS.TotalQty).setValue(Number(p.totalQty) || 0);
  }
  if (photo) renderPhotoCell(inv, row, INV_COLS.Photos, photo);
  if (p.notes !== undefined && p.notes !== '') {
    inv.getRange(row, INV_COLS.Notes).setValue(p.notes);
  }
  inv.getRange(row, INV_COLS.LastScanned).setValue(now);
  inv.getRange(row, INV_COLS.LastScannedBy).setValue(p.team || '');
  if (p.gpsLat) inv.getRange(row, INV_COLS.GpsLat).setValue(p.gpsLat);
  if (p.gpsLng) inv.getRange(row, INV_COLS.GpsLng).setValue(p.gpsLng);
  if (p.gpsAccuracy) inv.getRange(row, INV_COLS.GpsAcc).setValue(p.gpsAccuracy);

  appendScan(p, shelf, p.name, 'REGISTERED (updated)', p.notes || '');
  return { success: true, shelf: shelf, name: p.name, rowIdx: row, updated: true };
}

function actRegisterNoTag(p) {
  const inv = getSheet(CONFIG.SHEET_INVENTORY); ensureHeaders(inv, INV_HEADER);
  if (!p.name) return { success: false, error: 'Name is required' };

  const shelf = normShelf(p.shelf);
  if (shelf) return actRegister(p);  // shelf provided → normal path (with conflict protection)

  const now = nowDate();
  const photo = p.photo ? savePhoto(p.photo, p.name) : null;
  inv.appendRow([
    '', p.name || '', p.brand || '', p.type || '', 'in',
    Number(p.totalQty) || 0, 0, '',
    now, p.team || '', now, p.team || '',
    p.gpsLat || '', p.gpsLng || '', p.gpsAccuracy || '',
    p.notes || ''
  ]);
  const row = inv.getLastRow();
  if (photo) renderPhotoCell(inv, row, INV_COLS.Photos, photo);
  appendScan(p, '', p.name, 'REGISTERED (no shelf)', p.notes || '');
  return { success: true, name: p.name, rowIdx: row, created: true };
}

// Common ambiguity helper. If the caller didn't pin down a single row,
// surface the choices instead of guessing — the client will show a
// picker and re-submit with rowIdx baked in.
function ambiguousResponse(res, action) {
  return {
    success: false,
    ambiguous: true,
    error: 'Multiple items on shelf ' + (res.shelf || '?') + '. Pick which one to ' + action + '.',
    shelf: res.shelf,
    candidates: res.candidates
  };
}

function actTakeOut(p) {
  const inv = getSheet(CONFIG.SHEET_INVENTORY); ensureHeaders(inv, INV_HEADER);
  const res = resolveRow(inv, p);
  if (res.ambiguous) return ambiguousResponse(res, 'take out');
  if (res.row < 0) return { success: false, error: 'Item not found' };
  const row = res.row;
  const cur = readRow(inv, row);
  const qty = Number(p.qty) || 1;
  const total = Number(cur.totalQty) || 0;
  const newOut = Math.min(total || Infinity, (Number(cur.qtyOut) || 0) + qty);
  inv.getRange(row, INV_COLS.QtyOut).setValue(newOut);
  inv.getRange(row, INV_COLS.Status).setValue(total && newOut >= total ? 'out' : (newOut > 0 ? 'partial' : 'in'));
  inv.getRange(row, INV_COLS.LastScanned).setValue(nowDate());
  inv.getRange(row, INV_COLS.LastScannedBy).setValue(p.team || '');
  appendScan(p, cur.shelf, cur.name, 'TAKE OUT ×' + qty, p.reason || p.loadout || '');
  return { success: true, shelf: cur.shelf, name: cur.name, rowIdx: row, qtyOut: newOut, totalQty: total };
}

function actReturn(p) {
  const inv = getSheet(CONFIG.SHEET_INVENTORY); ensureHeaders(inv, INV_HEADER);
  const res = resolveRow(inv, p);
  if (res.ambiguous) return ambiguousResponse(res, 'return');
  if (res.row < 0) return { success: false, error: 'Item not found' };
  const row = res.row;
  const cur = readRow(inv, row);
  const qty = Number(p.qty) || 1;
  const newOut = Math.max(0, (Number(cur.qtyOut) || 0) - qty);
  inv.getRange(row, INV_COLS.QtyOut).setValue(newOut);
  inv.getRange(row, INV_COLS.Status).setValue(newOut === 0 ? 'in' : 'partial');
  inv.getRange(row, INV_COLS.LastScanned).setValue(nowDate());
  inv.getRange(row, INV_COLS.LastScannedBy).setValue(p.team || '');
  appendScan(p, cur.shelf, cur.name, 'RETURN ×' + qty, p.notes || '');
  return { success: true, shelf: cur.shelf, name: cur.name, rowIdx: row, qtyOut: newOut };
}

function actDamage(p) {
  const inv = getSheet(CONFIG.SHEET_INVENTORY); ensureHeaders(inv, INV_HEADER);
  const damage = getSheet(CONFIG.SHEET_DAMAGE); ensureHeaders(damage, DAMAGE_HEADER);
  // Damage doesn't strictly need a row to log against (the photo can
  // describe a generic shelf incident), but if we can resolve one we
  // include the proper item name in the log.
  const res = resolveRow(inv, p);
  if (res.ambiguous) return ambiguousResponse(res, 'log damage on');
  let shelf = '', name = p.name || '';
  if (res.row >= 2) {
    const cur = readRow(inv, res.row);
    shelf = cur.shelf; name = cur.name;
  } else {
    shelf = normShelf(p.shelf || p.id);
  }
  const photo = p.photo ? savePhoto(p.photo, name || shelf) : null;
  damage.appendRow([
    nowDate(), shelf || '', name || '', p.severity || 'minor', p.description || '',
    p.team || '', locLabel(p), p.gpsLat || '', p.gpsLng || '', ''
  ]);
  const dRow = damage.getLastRow();
  if (photo) renderPhotoCell(damage, dRow, DAMAGE_COL_PHOTOS, photo);
  appendScan(p, shelf, name, 'DAMAGE (' + (p.severity || 'minor') + ')', p.description || '');
  return { success: true, shelf: shelf, name: name };
}

function actUpdatePhoto(p) {
  const inv = getSheet(CONFIG.SHEET_INVENTORY); ensureHeaders(inv, INV_HEADER);
  const res = resolveRow(inv, p);
  if (res.ambiguous) return ambiguousResponse(res, 'update photo for');
  if (res.row < 0) return { success: false, error: 'Item not found' };
  const row = res.row;
  const cur = readRow(inv, row);
  if (!p.photo) return { success: false, error: 'No photo' };
  const photo = savePhoto(p.photo, cur.name || cur.shelf);
  renderPhotoCell(inv, row, INV_COLS.Photos, photo);
  inv.getRange(row, INV_COLS.LastScanned).setValue(nowDate());
  inv.getRange(row, INV_COLS.LastScannedBy).setValue(p.team || '');
  appendScan(p, cur.shelf, cur.name, 'PHOTO UPDATED', '');
  return { success: true, url: photo.url, rowIdx: row };
}

function actEdit(p) {
  const inv = getSheet(CONFIG.SHEET_INVENTORY); ensureHeaders(inv, INV_HEADER);
  const res = resolveRow(inv, p);
  if (res.ambiguous) return ambiguousResponse(res, 'edit');
  if (res.row < 0) return { success: false, error: 'Item not found' };
  const row = res.row;
  const cur = readRow(inv, row);
  // The Edit form may rename the item — if so, make sure we don't end
  // up with a duplicate (shelf, name) on the resulting shelf.
  const newName = (p.name !== undefined && p.name !== '') ? String(p.name).trim() : cur.name;
  const newShelf = p.newShelf ? normShelf(p.newShelf) : normShelf(cur.shelf);
  if ((newName !== cur.name || newShelf !== normShelf(cur.shelf))) {
    const clash = findRowByShelfAndName(inv, newShelf, newName);
    if (clash >= 0 && clash !== row) {
      return {
        success: false,
        error: '"' + newName + '" already exists on shelf ' + newShelf + '. Use a different name or shelf.',
        conflict: true
      };
    }
  }
  if (p.name !== undefined && p.name !== '')  inv.getRange(row, INV_COLS.Name).setValue(p.name);
  if (p.brand !== undefined && p.brand !== '') inv.getRange(row, INV_COLS.Brand).setValue(p.brand);
  if (p.type !== undefined && p.type !== '')  inv.getRange(row, INV_COLS.Type).setValue(p.type);
  if (p.totalQty !== undefined && p.totalQty !== '') {
    inv.getRange(row, INV_COLS.TotalQty).setValue(Number(p.totalQty) || 0);
  }
  if (p.newShelf) {
    const ns = normShelf(p.newShelf);
    if (ns) inv.getRange(row, INV_COLS.Shelf).setValue(ns);
  }
  if (p.notes !== undefined) {
    inv.getRange(row, INV_COLS.Notes).setValue(p.notes || '');
  }
  inv.getRange(row, INV_COLS.LastScanned).setValue(nowDate());
  inv.getRange(row, INV_COLS.LastScannedBy).setValue(p.team || '');
  appendScan(p, cur.shelf, p.name || cur.name, 'EDIT', p.notes || '');
  return { success: true, rowIdx: row };
}

function actLoad(p) {
  const inv = getSheet(CONFIG.SHEET_INVENTORY); ensureHeaders(inv, INV_HEADER);
  const res = resolveRow(inv, p);
  if (res.ambiguous) return ambiguousResponse(res, 'load');
  if (res.row < 0) return { success: false, error: 'Item not found' };
  const row = res.row;
  const cur = readRow(inv, row);
  const qty = Number(p.qty) || 1;
  const job = String(p.jobName || '').trim();
  if (!job) return { success: false, error: 'jobName required' };

  const loadSheetName = 'Load — ' + job;
  const spread = ss();
  let ls = spread.getSheetByName(loadSheetName);
  if (!ls) {
    ls = spread.insertSheet(loadSheetName);
    ls.getRange(1, 1, 1, 7).setValues([['Shelf','Item Name','Type','Qty Out','Taken By','Timestamp','Event']]).setFontWeight('bold');
    ls.setFrozenRows(1);
  }
  ls.appendRow([cur.shelf, cur.name, cur.type, qty, p.team || '', nowDate(), job]);

  const total = Number(cur.totalQty) || 0;
  const newOut = Math.min(total || Infinity, (Number(cur.qtyOut) || 0) + qty);
  inv.getRange(row, INV_COLS.QtyOut).setValue(newOut);
  inv.getRange(row, INV_COLS.Status).setValue(total && newOut >= total ? 'out' : (newOut > 0 ? 'partial' : 'in'));
  inv.getRange(row, INV_COLS.LastScanned).setValue(nowDate());
  inv.getRange(row, INV_COLS.LastScannedBy).setValue(p.team || '');
  appendScan(p, cur.shelf, cur.name, 'LOAD: ' + job + ' ×' + qty, job);
  return { success: true, loadSheet: loadSheetName, rowIdx: row };
}

function actDelivery(p) {
  const inv = getSheet(CONFIG.SHEET_INVENTORY); ensureHeaders(inv, INV_HEADER);
  // Delivery is usually a "scan a shelf as it moves" loop — we don't
  // require a single-item resolution. If only one item is on the shelf
  // we'll bump its lastScanned for traceability; otherwise we just log
  // the scan against the shelf with the team and direction.
  const shelf = normShelf(p.shelf || p.id);
  let name = '';
  if (shelf) {
    const rows = findRowsByShelf(inv, shelf);
    if (rows.length === 1) {
      name = readRow(inv, rows[0]).name;
      inv.getRange(rows[0], INV_COLS.LastScanned).setValue(nowDate());
      inv.getRange(rows[0], INV_COLS.LastScannedBy).setValue(p.team || '');
    }
  }
  appendScan(p, shelf, name, 'DELIVERY ' + (p.from || '?') + ' → ' + (p.to || '?'), p.note || '');
  return { success: true, shelf };
}

// ─────────── Setup / Maintenance ───────────
function setupSheets() {
  const sp = ss();
  const wanted = [
    [CONFIG.SHEET_INVENTORY, INV_HEADER],
    [CONFIG.SHEET_DAMAGE,    DAMAGE_HEADER],
    [CONFIG.SHEET_SCAN,      SCAN_HEADER]
  ];
  wanted.forEach(function (p) {
    let s = sp.getSheetByName(p[0]);
    if (!s) s = sp.insertSheet(p[0]);
    ensureHeaders(s, p[1]);
  });
  getOrCreateClientIdsSheet();
}

function setupDailyPrune() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'pruneClientIds') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('pruneClientIds').timeBased().everyDays(1).atHour(3).create();
}
