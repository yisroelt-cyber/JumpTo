// src/services/jumpToStorage.js
/* global Excel, OfficeRuntime */

import { MAX_RECENTS } from "../shared/constants";
const SETTINGS_SHEET_NAME = "_JumpToAddinSettings";
const USERKEY_STORAGE_KEY = "JumpTo.UserKey";

export const MAX_FAVORITES = 20;
// Row indices (1-based)
const ROW_USERKEY = 1;
const ROW_FAVORITES = 2;
const ROW_RECENTS = 3;
const ROW_SETTINGS = 4;

// Inventory table start (1-based)
const INV_START_ROW = 52;


// OfficeRuntime.storage — LRU-10 safety-net cache (Favorites + Settings only)
// Identity is keyed by (workbookGuid + filenameFingerprint). We encode this into a compact key
// using workbookGuid plus a short hash of filenameFingerprint.
const RT10_INDEX_KEY = "JumpTo.RT10.Index";
const RT10_ENTRY_PREFIX = "JumpTo.RT10.Entry.";
const RT10_MAX = 10;

function hashStringToBase36(s) {
  // Stable, fast, non-cryptographic hash (djb2) -> unsigned 32-bit -> base36
  let h = 5381;
  const str = String(s || "");
  for (let i = 0; i < str.length; i++) {
    h = ((h << 5) + h) + str.charCodeAt(i); // h*33 + c
    h = h >>> 0;
  }
  return h.toString(36);
}

function makeRt10EntryKey(workbookGuid, filenameFingerprint) {
  const g = String(workbookGuid || "").trim();
  const f = String(filenameFingerprint || "").trim();
  return `${RT10_ENTRY_PREFIX}${g}.${hashStringToBase36(f)}`;
}

async function rt10GetIndex() {
  try {
    const raw = await OfficeRuntime.storage.getItem(RT10_INDEX_KEY);
    const idx = safeJsonParse(raw, { v: 1, entries: [] });
    if (!idx || typeof idx !== "object" || !Array.isArray(idx.entries)) return { v: 1, entries: [] };
    return idx;
  } catch {
    return { v: 1, entries: [] };
  }
}

async function rt10SetIndex(idx) {
  try {
    await OfficeRuntime.storage.setItem(RT10_INDEX_KEY, safeJsonStringify(idx));
  } catch {}
}

async function rt10Touch(key) {
  const idx = await rt10GetIndex();
  const now = Date.now();
  const entries = idx.entries.filter(e => e && e.key && e.key !== key);
  entries.unshift({ key, lastAccess: now });
  idx.entries = entries.slice(0, RT10_MAX);
  await rt10SetIndex(idx);
}

async function rt10EvictOverflow() {
  const idx = await rt10GetIndex();
  if (idx.entries.length <= RT10_MAX) return;
  const keep = idx.entries.slice(0, RT10_MAX);
  const evict = idx.entries.slice(RT10_MAX);
  idx.entries = keep;
  await rt10SetIndex(idx);
  for (const e of evict) {
    try { await OfficeRuntime.storage.removeItem(e.key); } catch {}
  }
}

async function rt10Read(workbookGuid, filenameFingerprint) {
  if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage?.getItem) return null;
  const key = makeRt10EntryKey(workbookGuid, filenameFingerprint);
  try {
    const raw = await OfficeRuntime.storage.getItem(key);
    if (!raw) return null;
    const obj = safeJsonParse(raw, null);
    if (!obj || typeof obj !== "object") return null;
    // Validate identity match
    if (obj.workbookGuid !== workbookGuid) return null;
    if (obj.filenameFingerprint !== filenameFingerprint) return null;
    await rt10Touch(key);
    return obj;
  } catch {
    return null;
  }
}

async function rt10Write(workbookGuid, filenameFingerprint, favorites, settings) {
  if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage?.setItem) return;
  const key = makeRt10EntryKey(workbookGuid, filenameFingerprint);
  const payload = {
    workbookGuid,
    filenameFingerprint,
    dts: Date.now(),
    favorites: Array.isArray(favorites) ? favorites : [],
    settings: (settings && typeof settings === "object") ? settings : {}
  };
  try {
    await OfficeRuntime.storage.setItem(key, safeJsonStringify(payload));
  } catch {}
  await rt10Touch(key);
  await rt10EvictOverflow();
}

function safeJsonParse(str, fallback) {
  try {
    if (typeof str !== "string") return fallback;
    const s = str.trim();
    if (!s) return fallback;
    return JSON.parse(s);
  } catch {
    return fallback;
  }
}

function safeJsonStringify(obj) {
  try {
    return JSON.stringify(obj);
  } catch {
    return "[]";
  }
}

function isPlainObjectEmpty(o) {
  if (!o || typeof o !== "object" || Array.isArray(o)) return true;
  return Object.keys(o).length === 0;
}


async function getOrCreateUserKey() {
  // Prefer OfficeRuntime.storage, but fall back to Office.context.roamingSettings if storage is unavailable
  // or not persisting across sessions in this host.
  let existing = null;

  // 1) OfficeRuntime.storage (Shared Runtime)
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage?.getItem) {
      existing = await OfficeRuntime.storage.getItem(USERKEY_STORAGE_KEY);
      if (existing) return existing;
    }
  } catch {}

  // 2) Roaming settings (per-user, persists across sessions)
  try {
    const rs = Office?.context?.roamingSettings;
    if (rs?.get) {
      existing = rs.get(USERKEY_STORAGE_KEY);
      if (existing) return existing;
    }
  } catch {}

  // 3) localStorage (last resort; persists per-browser)
  try {
    existing = globalThis?.localStorage?.getItem?.(USERKEY_STORAGE_KEY);
    if (existing) return existing;
  } catch {}

  // Create new key
  const key =
    (globalThis.crypto?.randomUUID?.() ||
      `u_${Date.now()}_${Math.random().toString(16).slice(2)}`);

  // Persist to all backends that are available.
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage?.setItem) {
      await OfficeRuntime.storage.setItem(USERKEY_STORAGE_KEY, key);
    }
  } catch {}

  try {
    const rs = Office?.context?.roamingSettings;
    if (rs?.set && rs?.saveAsync) {
      rs.set(USERKEY_STORAGE_KEY, key);
      await new Promise((resolve) => rs.saveAsync(() => resolve()));
    }
  } catch {}

  try {
    globalThis?.localStorage?.setItem?.(USERKEY_STORAGE_KEY, key);
  } catch {}

  return key;
}

function colIndexToLetter(idx1) {
  // 1-based index to A1 letter(s)
  let idx = idx1;
  let s = "";
  while (idx > 0) {
    const r = (idx - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    idx = Math.floor((idx - 1) / 26);
  }
  return s;
}

async function ensureSettingsSheet(context) {
  const ws = context.workbook.worksheets.getItemOrNullObject(SETTINGS_SHEET_NAME);
  ws.load("name,visibility");
  await context.sync();

  if (!ws.isNullObject) {
    // Enforce invariant: JumpTo settings sheet is always VeryHidden.
    if (ws.visibility !== Excel.SheetVisibility.veryHidden) {
      ws.visibility = Excel.SheetVisibility.veryHidden;
      await context.sync();
    }
    return ws;
  }

  const created = context.workbook.worksheets.add(SETTINGS_SHEET_NAME);
  created.visibility = Excel.SheetVisibility.veryHidden;
  created.load("name");
  await context.sync();
  return created;
}

// --- Workbook identity (Phase 1 groundwork) ---
// Workbook GUID is stored inside the VeryHidden JumpTo settings sheet.
// Filename fingerprint is stored as the formula =CELL("filename") in the same sheet.
//
// Storage location (reserved):
//   A1: label "JT_GUID"        B1: GUID value
//   A2: label "JT_FILENAME"    B2: formula =CELL("filename") (value read as fingerprint)
const WB_ID_RANGE_ADDRESS = "A1:B2";
const WB_GUID_LABEL = "JT_GUID";
const WB_FILENAME_LABEL = "JT_FILENAME";

function isValidGuid(s) {
  return typeof s === "string" && /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(s.trim());
}

async function ensureWorkbookIdentity(context, settingsSheet) {
  const range = settingsSheet.getRange(WB_ID_RANGE_ADDRESS);
  range.load("values,formulas");
  await context.sync();

  const values = range.values || [["", ""], ["", ""]];
  const a1 = (values[0]?.[0] || "").toString();
  const b1 = (values[0]?.[1] || "").toString();
  const a2 = (values[1]?.[0] || "").toString();
  const b2 = (values[1]?.[1] || "").toString();

  let guid = b1 && isValidGuid(b1) ? b1.trim() : null;

  // Ensure labels exist (best-effort, harmless if already present)
  if (a1 !== WB_GUID_LABEL) range.getCell(0, 0).values = [[WB_GUID_LABEL]];
  if (a2 !== WB_FILENAME_LABEL) range.getCell(1, 0).values = [[WB_FILENAME_LABEL]];

  // Ensure GUID exists
  if (!guid) {
    guid = (typeof crypto !== "undefined" && crypto.randomUUID) ? crypto.randomUUID() : (
      // Fallback UUID v4 generator (no dependencies)
      "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
        const r = Math.random() * 16 | 0;
        const v = c === "x" ? r : (r & 0x3 | 0x8);
        return v.toString(16);
      })
    );
    range.getCell(0, 1).values = [[guid]];
  }

  // Ensure filename fingerprint formula exists
  // Note: value may be blank if workbook isn't saved yet; that is acceptable for our safety-net design.
  const cellB2 = range.getCell(1, 1);
  const existingFormula = (range.formulas?.[1]?.[1] || "").toString();
  if (!existingFormula || existingFormula.toUpperCase() !== '=CELL("FILENAME")') {
    cellB2.formulas = [[`=CELL("filename")`]];
  }

  await context.sync();

  // Read the computed fingerprint value
  cellB2.load("values");
  await context.sync();
  const filenameFingerprint = (cellB2.values?.[0]?.[0] || "").toString();

  return { workbookGuid: guid, filenameFingerprint };
}

export async function getWorkbookIdentity() {
  return Excel.run(async (context) => {
    const settingsSheet = await ensureSettingsSheet(context);
    return ensureWorkbookIdentity(context, settingsSheet);
  });
}

async function getUserColumn(context, settingsSheet, userKey) {
  // Search row 1 starting from column D for userKey; else first empty.
  // We'll scan D1:ZZ1 (~700 cols) which is plenty.
  const headerRange = settingsSheet.getRange("D1:ZZ1");
  headerRange.load("values");
  await context.sync();

  const values = headerRange.values?.[0] || [];
  let foundOffset = -1;
  let emptyOffset = -1;
  for (let i = 0; i < values.length; i++) {
    const v = values[i];
    if (v === userKey) { foundOffset = i; break; }
    if (emptyOffset === -1 && (v === null || v === "")) emptyOffset = i;
  }

  const offset = foundOffset !== -1 ? foundOffset : (emptyOffset !== -1 ? emptyOffset : values.length);
  // D is column 4
  const colIdx1 = 4 + offset;
  const colLetter = colIndexToLetter(colIdx1);

  if (foundOffset === -1) {
    // write userKey into ROW_USERKEY at this column
    const cell = settingsSheet.getRange(`${colLetter}${ROW_USERKEY}`);
    cell.values = [[userKey]];
    await context.sync();
  }

  return { colIdx1, colLetter };
}

async function readUserCells(context, sheet, colLetter) {
  const favCell = sheet.getRange(`${colLetter}${ROW_FAVORITES}`);
  const recCell = sheet.getRange(`${colLetter}${ROW_RECENTS}`);
  const setCell = sheet.getRange(`${colLetter}${ROW_SETTINGS}`);
  favCell.load("values");
  recCell.load("values");
  setCell.load("values");
  await context.sync();

  const favorites = safeJsonParse(favCell.values?.[0]?.[0], []);
  const recents = safeJsonParse(recCell.values?.[0]?.[0], []);
  const settings = safeJsonParse(setCell.values?.[0]?.[0], {});

  return { favorites, recents, settings };
}

async function writeUserCells(context, sheet, colLetter, { favorites, recents, settings }) {
  const favCell = sheet.getRange(`${colLetter}${ROW_FAVORITES}`);
  const recCell = sheet.getRange(`${colLetter}${ROW_RECENTS}`);
  const setCell = sheet.getRange(`${colLetter}${ROW_SETTINGS}`);
  favCell.values = [[safeJsonStringify(favorites || [])]];
  recCell.values = [[safeJsonStringify(recents || [])]];
  setCell.values = [[safeJsonStringify(settings || {})]];
  await context.sync();
}

async function loadInventory(context, sheet, userColLetter) {
  // Load A52:C2000 and user's frequency column range for same rows.
  // Column A = id, B = name, C reserved (blank). User col stores frequency.
  const endRow = 2000;
  const invRange = sheet.getRange(`A${INV_START_ROW}:C${endRow}`);
  const freqRange = sheet.getRange(`${userColLetter}${INV_START_ROW}:${userColLetter}${endRow}`);
  invRange.load("values");
  freqRange.load("values");
  await context.sync();

  const inv = invRange.values || [];
  const freq = freqRange.values || [];
  const rows = [];
  for (let i = 0; i < inv.length; i++) {
    const rowNum = INV_START_ROW + i;
    const id = inv[i]?.[0] ?? "";
    const name = inv[i]?.[1] ?? "";
    const f = freq[i]?.[0];
    rows.push({ rowNum, id: String(id || ""), name: String(name || ""), freq: typeof f === "number" ? f : Number(f || 0) });
  }
  return rows;
}

async function syncInventoryWithVisibleSheets(context, sheet, userColLetter, visibleSheets) {
  // visibleSheets: [{id,name,orderIndex?}]
  const rows = await loadInventory(context, sheet, userColLetter);

  // Build maps of existing rows
  const idToRow = new Map();
  const nameToRow = new Map();
  let lastUsedRow = INV_START_ROW - 1;

  for (const r of rows) {
    if (r.id || r.name) lastUsedRow = r.rowNum;
    if (r.id) idToRow.set(r.id, r.rowNum);
    if (r.name) nameToRow.set(r.name, r.rowNum);
  }

  const matchedRows = new Set();

  // Assign or update rows for each visible sheet
  for (const s of visibleSheets) {
    const sid = String(s.id || "");
    const sname = String(s.name || "");
    if (!sid || !sname) continue;

    let rowNum = idToRow.get(sid);
    if (rowNum) {
      // update name if needed
      const nameCell = sheet.getRange(`B${rowNum}`);
      nameCell.values = [[sname]];
      matchedRows.add(rowNum);
      continue;
    }

    rowNum = nameToRow.get(sname);
    if (rowNum) {
      // update id if needed
      const idCell = sheet.getRange(`A${rowNum}`);
      idCell.values = [[sid]];
      matchedRows.add(rowNum);
      continue;
    }

    // append new
    lastUsedRow += 1;
    const idCell = sheet.getRange(`A${lastUsedRow}`);
    const nameCell = sheet.getRange(`B${lastUsedRow}`);
    idCell.values = [[sid]];
    nameCell.values = [[sname]];
    // initialize freq to 0
    const fCell = sheet.getRange(`${userColLetter}${lastUsedRow}`);
    fCell.values = [[0]];
    matchedRows.add(lastUsedRow);
  }

  // Clear rows that are not matched but contain data
  for (const r of rows) {
    if ((r.id || r.name) && !matchedRows.has(r.rowNum)) {
      sheet.getRange(`A${r.rowNum}:C${r.rowNum}`).clear();
      sheet.getRange(`${userColLetter}${r.rowNum}`).clear();
    }
  }

  await context.sync();
}

async function incrementFrequency(context, sheet, userColLetter, sheetId) {
  const rows = await loadInventory(context, sheet, userColLetter);
  const target = rows.find(r => r.id === sheetId) || null;
  if (!target) return 0;

  const cell = sheet.getRange(`${userColLetter}${target.rowNum}`);
  cell.load("values");
  await context.sync();

  const cur = Number(cell.values?.[0]?.[0] || 0);
  const next = cur + 1;
  cell.values = [[next]];
  await context.sync();
  return next;
}

export async function getJumpToState() {
  const userKey = await getOrCreateUserKey();
  if (!userKey) {
    return { userKey: null, sheets: [], favorites: [], recents: [], settings: {}, global: {} };
  }

  // Read workbook state first (includes ensuring workbook identity fields exist).
  const wb = await Excel.run(async (context) => {
    const settingsSheet = await ensureSettingsSheet(context);
    const wbId = await ensureWorkbookIdentity(context, settingsSheet);
    const { colLetter } = await getUserColumn(context, settingsSheet, userKey);

    // Load visible sheets with id+name and workbook order
    const sheets = context.workbook.worksheets;
    sheets.load("items/id,name,visibility");
    await context.sync();
    const visible = sheets.items.filter(ws => ws.visibility === Excel.SheetVisibility.visible);
    const visibleSheets = visible.map((ws, idx) => ({ id: ws.id, name: ws.name, orderIndex: idx }));

    // Reconcile inventory and read per-user blobs
    await syncInventoryWithVisibleSheets(context, settingsSheet, colLetter, visibleSheets);
    const { favorites, recents, settings } = await readUserCells(context, settingsSheet, colLetter);

    // Build enriched favorites/recents objects with names
    const idToName = new Map(visibleSheets.map(s => [s.id, s.name]));
    const favObjs = (Array.isArray(favorites) ? favorites : []).map(id => ({ id, name: idToName.get(id) || "" }));
    const recObjs = (Array.isArray(recents) ? recents : []).map(id => ({ id, name: idToName.get(id) || "" }));

    // Load frequency values for visible sheets (for ordering)
    const invRows = await loadInventory(context, settingsSheet, colLetter);
    const freqById = {};
    for (const r of invRows) {
      if (r.id) freqById[r.id] = Number(r.freq || 0);
    }

    return {
      __wbId: wbId,
      userKey,
      sheets: visibleSheets,
      favorites: favObjs,
      recents: recObjs,
      settings: (settings && typeof settings === "object") ? settings : {},
      global: { freqById }
    };
  });

  // Phase 2: runtime safety net (Favorites + Settings only) gated by (GUID + filename fingerprint).
  const { workbookGuid, filenameFingerprint } = wb.__wbId || {};
  let favorites = wb.favorites;
  let settings = wb.settings;

  if (workbookGuid && filenameFingerprint) {
    const rt = await rt10Read(workbookGuid, filenameFingerprint);

    // If workbook lost state (common failure mode), prefer runtime non-empty values.
    if (rt) {
      if ((!Array.isArray(favorites) || favorites.length === 0) && Array.isArray(rt.favorites) && rt.favorites.length > 0) {
        const idToName = new Map((wb.sheets || []).map(s => [s.id, s.name]));
        favorites = rt.favorites.map(id => ({ id, name: idToName.get(id) || "" }));
      }
      if (isPlainObjectEmpty(settings) && rt.settings && typeof rt.settings === "object" && !isPlainObjectEmpty(rt.settings)) {
        settings = rt.settings;
      }
    } else {
      // Seed runtime safety net from workbook if there is meaningful data.
      const favIds = (Array.isArray(favorites) ? favorites : []).map(f => f?.id).filter(Boolean);
      if (favIds.length > 0 || (settings && typeof settings === "object" && !isPlainObjectEmpty(settings))) {
        await rt10Write(workbookGuid, filenameFingerprint, favIds, settings);
      }
    }
  }

  return { ...wb, favorites, settings };
}


export async function toggleFavorite(sheetId) {
  const userKey = await getOrCreateUserKey();
  if (!userKey) return null;

  return Excel.run(async (context) => {
    const sheet = await ensureSettingsSheet(context);
    const { colLetter } = await getUserColumn(context, sheet, userKey);
    const state = await readUserCells(context, sheet, colLetter);
    const favs = Array.isArray(state.favorites) ? [...state.favorites] : [];

    const idx = favs.indexOf(sheetId);
    if (idx >= 0) {
      favs.splice(idx, 1);
    } else {
      favs.push(sheetId);
      if (favs.length > MAX_FAVORITES) favs.length = MAX_FAVORITES;
    }

    await writeUserCells(context, sheet, colLetter, { favorites: favs, recents: state.recents, settings: state.settings });
    return favs;
  });
}


export async function setFavorites(favIds) {
  const userKey = await getOrCreateUserKey();
  if (!userKey) return;

  const nextFavs = Array.isArray(favIds) ? favIds.filter(Boolean) : [];

  const out = await Excel.run(async (context) => {
    const settingsSheet = await ensureSettingsSheet(context);
    const wbId = await ensureWorkbookIdentity(context, settingsSheet);
    const { colLetter } = await getUserColumn(context, settingsSheet, userKey);
    const { recents, settings } = await readUserCells(context, settingsSheet, colLetter);
    await writeUserCells(context, settingsSheet, colLetter, { favorites: nextFavs, recents, settings });
    return { wbId, settings };
  });

  const { workbookGuid, filenameFingerprint } = out?.wbId || {};
  if (workbookGuid && filenameFingerprint) {
    await rt10Write(workbookGuid, filenameFingerprint, nextFavs, out.settings || {});
  }
}

export async function addFavorite(sheetId) {
  if (!sheetId) return null;
  const state = await getJumpToState();
  const current = Array.isArray(state.favorites) ? state.favorites.map(x => x?.id).filter(Boolean) : [];
  if (current.includes(sheetId)) return current;
  const next = [...current, sheetId].slice(0, MAX_FAVORITES);
  return setFavorites(next);
}

export async function removeFavorite(sheetId) {
  if (!sheetId) return null;
  const state = await getJumpToState();
  const current = Array.isArray(state.favorites) ? state.favorites.map(x => x?.id).filter(Boolean) : [];
  const next = current.filter(id => id !== sheetId);
  return setFavorites(next);
}

export async function moveFavorite(sheetId, direction) {
  if (!sheetId) return null;
  if (direction !== "up" && direction !== "down") return null;

  const state = await getJumpToState();
  const current = Array.isArray(state.favorites) ? state.favorites.map(x => x?.id).filter(Boolean) : [];
  const idx = current.indexOf(sheetId);
  if (idx < 0) return current;

  const to = direction === "up" ? idx - 1 : idx + 1;
  if (to < 0 || to >= current.length) return current;

  const next = current.slice();
  const [item] = next.splice(idx, 1);
  next.splice(to, 0, item);
  return setFavorites(next);
}




export async function setUiSettings(settingsPatch) {
  const userKey = await getOrCreateUserKey();
  if (!userKey) return;

  const patch = (settingsPatch && typeof settingsPatch === "object") ? settingsPatch : {};

  const out = await Excel.run(async (context) => {
    const settingsSheet = await ensureSettingsSheet(context);
    const wbId = await ensureWorkbookIdentity(context, settingsSheet);
    const { colLetter } = await getUserColumn(context, settingsSheet, userKey);

    const { favorites, recents, settings } = await readUserCells(context, settingsSheet, colLetter);
    const nextSettings = { ...(settings || {}), ...patch };

    await writeUserCells(context, settingsSheet, colLetter, { favorites, recents, settings: nextSettings });

    return { wbId, favorites, nextSettings };
  });

  const { workbookGuid, filenameFingerprint } = out?.wbId || {};
  if (workbookGuid && filenameFingerprint) {
    await rt10Write(workbookGuid, filenameFingerprint, out.favorites || [], out.nextSettings || {});
  }
}

export async function recordActivation(sheetId) {
  const userKey = await getOrCreateUserKey();
  if (!userKey) return null;

  return Excel.run(async (context) => {
    const settingsSheet = await ensureSettingsSheet(context);
    const { colLetter } = await getUserColumn(context, settingsSheet, userKey);

    // Get current visible sheets (for reconciliation)
    const ws = context.workbook.worksheets;
    ws.load("items/id,name,visibility");
    await context.sync();
    const visible = ws.items.filter(w => w.visibility === Excel.SheetVisibility.visible);
    const visibleSheets = visible.map((w, idx) => ({ id: w.id, name: w.name, orderIndex: idx }));

    await syncInventoryWithVisibleSheets(context, settingsSheet, colLetter, visibleSheets);

    const state = await readUserCells(context, settingsSheet, colLetter);

    // Update recents
    const rec = Array.isArray(state.recents) ? [...state.recents] : [];
    const existing = rec.indexOf(sheetId);
    if (existing >= 0) rec.splice(existing, 1);
    rec.unshift(sheetId);
    if (rec.length > MAX_RECENTS) rec.length = MAX_RECENTS;

    await writeUserCells(context, settingsSheet, colLetter, { favorites: state.favorites, recents: rec, settings: state.settings });
    const nextFreq = await incrementFrequency(context, settingsSheet, colLetter, sheetId);

    return { recents: rec, freq: nextFreq };
  });
}