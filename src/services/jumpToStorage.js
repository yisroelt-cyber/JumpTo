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
  ws.load("name");
  await context.sync();
  if (!ws.isNullObject) return ws;

  const created = context.workbook.worksheets.add(SETTINGS_SHEET_NAME);
  created.visibility = Excel.SheetVisibility.veryHidden;
  // Make sure it stays veryHidden even if user toggles sheet visibility UI
  created.load("name");
  await context.sync();
  return created;
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

  return Excel.run(async (context) => {
    const settingsSheet = await ensureSettingsSheet(context);
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

    // Global options
    let global = {};
    try {
      const oneDigit = await OfficeRuntime.storage.getItem("JumpTo.Option.OneDigitActivation");
      const rowHeight = await OfficeRuntime.storage.getItem("JumpTo.Option.RowHeightPreset");
      const baseline = await OfficeRuntime.storage.getItem("JumpTo.Option.BaselineOrder"); // "workbook" | "alpha"
      const freqOnTop = await OfficeRuntime.storage.getItem("JumpTo.Option.FrequentOnTop");
      global = {
        oneDigitActivationEnabled: oneDigit !== "false",
        rowHeightPreset: rowHeight || "Compact",
        baselineOrder: baseline || "workbook",
        frequentOnTop: freqOnTop !== "false",
      };
    } catch {
      global = { oneDigitActivationEnabled: true, rowHeightPreset: "Compact", baselineOrder: "workbook", frequentOnTop: true };
    }

    return {
      userKey,
      sheets: visibleSheets.map(s => ({ ...s, freq: freqById[s.id] || 0 })),
      favorites: favObjs,
      recents: recObjs,
      settings: settings || {},
      global,
    };
  });
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


export async function setFavorites(favoriteIds) {
  const userKey = await getOrCreateUserKey();
  if (!userKey) return null;

  const ids = Array.isArray(favoriteIds) ? favoriteIds.filter(Boolean).slice(0, MAX_FAVORITES) : [];

  return Excel.run(async (context) => {
    const sheet = await ensureSettingsSheet(context);
    const { colLetter } = await getUserColumn(context, sheet, userKey);
    const state = await readUserCells(context, sheet, colLetter);

    // Reduce perceived lag: suppress UI work & calc for this sync (without changing global calc mode)
    try {
      context.workbook.application.suspendScreenUpdatingUntilNextSync();
      context.workbook.application.suspendApiCalculationUntilNextSync();
    } catch {
      // ignore if host doesn't support
    }

    await writeUserCells(context, sheet, colLetter, { favorites: ids, recents: state.recents, settings: state.settings });
    return ids;
  });
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

  return Excel.run(async (context) => {
    const settingsSheet = await ensureSettingsSheet(context);
    const { colLetter } = await getUserColumn(context, settingsSheet, userKey);

    const { favorites, recents, settings } = await readUserCells(context, settingsSheet, colLetter);
    const nextSettings = { ...(settings || {}), ...patch };

    await writeUserCells(context, settingsSheet, colLetter, { favorites, recents, settings: nextSettings });
  });
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