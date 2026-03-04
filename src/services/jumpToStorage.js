// 2026-03-04 21:15 UTC
/* global Excel, OfficeRuntime */

import { MAX_RECENTS, MAX_FAVORITES } from "../shared/constants";
const SETTINGS_SHEET_NAME = "_JumpToAddinSettings";
const USERKEY_STORAGE_KEY = "JumpTo.UserKey";

// Row indices (1-based)
const ROW_USERKEY = 1;
const ROW_FAVORITES = 2;
const ROW_RECENTS = 3;
const ROW_SETTINGS = 4;

// Dev premium flag cell: write "DEV_PREMIUM" here to enable frequency bump for testing.
const DEV_FLAG_CELL = "A10";
const DEV_FLAG_VALUE = "DEV_PREMIUM";

// Workbook-scoped user settings persisted in the workbook Settings sheet blob.
const WB_SETTINGS_KEYS = ["oneDigitActivationEnabled"];

function pickWbSettings(obj) {
  const src = (obj && typeof obj === "object" && !Array.isArray(obj)) ? obj : {};
  const out = {};
  for (const k of WB_SETTINGS_KEYS) {
    if (Object.prototype.hasOwnProperty.call(src, k)) out[k] = src[k];
  }
  return out;
}


// Diagnostics: identity logging disabled in production.
function identityLog(_message, _data) {
  // no-op
}


// Inventory table start (1-based)
const INV_START_ROW = 52;


// OfficeRuntime.storage — LRU-10 safety-net cache (Favorites + Settings only)
// Identity is keyed by (workbookGuid + filenameFingerprint). We encode this into a compact key
// using workbookGuid plus a short hash of filenameFingerprint.
const RT10_INDEX_KEY = "JumpTo.RT10.Index";
const RT10_ENTRY_PREFIX = "JumpTo.RT10.Entry.";
const RT10_MAX = 10;

// OfficeRuntime.storage — LRU-3 performance cache (Visible worksheets + Recents)
// Cache-only: bounded, allowed to be stale, never authoritative.
// Identity is keyed by (workbookGuid + filenameFingerprint), same as RT10.
const RT3_INDEX_KEY = "JumpTo.RT3.Index";
const RT3_ENTRY_PREFIX = "JumpTo.RT3.Entry.";
const RT3_MAX = 3;


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


function makeRt3EntryKey(workbookGuid, filenameFingerprint) {
  const g = String(workbookGuid || "").trim();
  const f = String(filenameFingerprint || "").trim();
  return `${RT3_ENTRY_PREFIX}${g}.${hashStringToBase36(f)}`;
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


async function rt3GetIndex() {
  try {
    const raw = await OfficeRuntime.storage.getItem(RT3_INDEX_KEY);
    const parsed = raw ? JSON.parse(raw) : null;
    if (parsed && Array.isArray(parsed.entries)) return parsed;
  } catch {}
  return { entries: [] };
}

async function rt3SetIndex(idx) {
  try { await OfficeRuntime.storage.setItem(RT3_INDEX_KEY, JSON.stringify(idx)); } catch {}
}

async function rt3Touch(entryKey) {
  const idx = await rt3GetIndex();
  const now = Date.now();
  const next = idx.entries.filter(e => e.key !== entryKey);
  next.unshift({ key: entryKey, ts: now });
  idx.entries = next;
  await rt3SetIndex(idx);
  await rt3EvictOverflow();
}

async function rt3EvictOverflow() {
  const idx = await rt3GetIndex();
  if (idx.entries.length <= RT3_MAX) return;
  const keep = idx.entries.slice(0, RT3_MAX);
  const evict = idx.entries.slice(RT3_MAX);
  idx.entries = keep;
  await rt3SetIndex(idx);
  for (const e of evict) {
    try { await OfficeRuntime.storage.removeItem(e.key); } catch {}
  }
}

async function rt3Read(workbookGuid, filenameFingerprint) {
  if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage?.getItem) return null;
  const key = makeRt3EntryKey(workbookGuid, filenameFingerprint);
  try {
    const raw = await OfficeRuntime.storage.getItem(key);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") return null;
    await rt3Touch(key);
    return parsed;
  } catch {
    return null;
  }
}

async function rt3Write(workbookGuid, filenameFingerprint, payload) {
  if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage?.setItem) return;
  const key = makeRt3EntryKey(workbookGuid, filenameFingerprint);
  try {
    await OfficeRuntime.storage.setItem(key, JSON.stringify(payload || {}));
  } catch {
    // ignore
  }
  await rt3Touch(key);
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

async function rt10Write(workbookGuid, filenameFingerprint, favorites, settings, dtsOverride) {
  if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage?.setItem) return;
  const key = makeRt10EntryKey(workbookGuid, filenameFingerprint);
  const payload = {
    workbookGuid,
    filenameFingerprint,
    dts: (typeof dtsOverride === "number" ? dtsOverride : Date.now()),
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


// --- Frequency decay ---
// Storage format: JSON string { freq: <float>, dts: <ms timestamp> }
// Half-life: 30 days. Decay is applied on both write and read so that
// sheets not activated recently fade out naturally.

const FREQ_HALF_LIFE_MS = 30 * 24 * 60 * 60 * 1000; // 30 days in ms

function decayFreq(freq, dts, nowMs) {
  const f = typeof freq === "number" && isFinite(freq) ? freq : 0;
  const d = typeof dts === "number" && isFinite(dts) && dts > 0 ? dts : null;
  if (!d || f <= 0) return f;
  const elapsed = Math.max(0, nowMs - d);
  return f * Math.pow(0.5, elapsed / FREQ_HALF_LIFE_MS);
}

function parseFreqCell(raw) {
  // Accept new format: JSON string { freq, dts }
  // Gracefully handle legacy plain numbers: treat as { freq: 0, dts: 0 }
  // so old test records decay to nothing immediately.
  try {
    if (raw === null || raw === undefined || raw === "") return { freq: 0, dts: 0 };
    if (typeof raw === "number") return { freq: 0, dts: 0 };
    if (typeof raw === "string") {
      const s = raw.trim();
      if (!s) return { freq: 0, dts: 0 };
      // Plain numeric string = legacy
      if (/^-?\d+(\.\d+)?$/.test(s)) return { freq: 0, dts: 0 };
      const v = JSON.parse(s);
      if (v && typeof v === "object" && !Array.isArray(v)) {
        const freq = typeof v.freq === "number" && isFinite(v.freq) ? v.freq : 0;
        const dts = typeof v.dts === "number" && isFinite(v.dts) ? v.dts : 0;
        return { freq, dts };
      }
    }
  } catch {
    // ignore
  }
  return { freq: 0, dts: 0 };
}

function stringifyFreqCell(freq, dts) {
  return JSON.stringify({ freq, dts });
}


// Normalize a raw favorites entry to a FavoriteItem object.
// Accepts either a plain string ID (legacy) or a partial/full object.
function normalizeFavoriteItem(entry, index) {
  if (typeof entry === "string" && entry) {
    // Legacy plain-ID — derive digit and displayOrder from position
    return { id: entry, workbookId: "this", digit: index + 1, displayOrder: index + 1, name: "" };
  }
  if (entry && typeof entry === "object" && entry.id) {
    return {
      id:           String(entry.id),
      workbookId:   typeof entry.workbookId === "string" ? entry.workbookId : "this",
      digit:        typeof entry.digit === "number" ? entry.digit : index + 1,
      displayOrder: typeof entry.displayOrder === "number" ? entry.displayOrder : index + 1,
      name:         typeof entry.name === "string" ? entry.name : ""
    };
  }
  return null;
}

// Normalize a raw recents entry to a RecentItem object.
function normalizeRecentItem(entry) {
  if (typeof entry === "string" && entry) {
    return { id: entry, workbookId: "this", name: "" };
  }
  if (entry && typeof entry === "object" && entry.id) {
    return {
      id:         String(entry.id),
      workbookId: typeof entry.workbookId === "string" ? entry.workbookId : "this",
      name:       typeof entry.name === "string" ? entry.name : ""
    };
  }
  return null;
}

function parseFavoritesCell(raw) {
  // Accepts:
  //   Legacy bare array of plain IDs:   ["id1","id2",...]
  //   Legacy wrapped plain IDs:          { dts: ms, favorites: ["id1","id2",...] }
  //   New wrapped object array:          { dts: ms, favorites: [{id,workbookId,digit,displayOrder,name},...] }
  try {
    if (typeof raw !== "string") raw = String(raw ?? "");
    const s = raw.trim();
    if (!s) return { favorites: [], dts: 0, valid: false };
    const v = JSON.parse(s);

    // Bare array (legacy)
    if (Array.isArray(v)) {
      const items = v.map((e, i) => normalizeFavoriteItem(e, i)).filter(Boolean);
      return { favorites: items, dts: 0, valid: true };
    }

    // Wrapped object { dts, favorites: [...] }
    if (v && typeof v === "object" && Array.isArray(v.favorites)) {
      const dts = (typeof v.dts === "number" && isFinite(v.dts)) ? v.dts : 0;
      const items = v.favorites.map((e, i) => normalizeFavoriteItem(e, i)).filter(Boolean);
      return { favorites: items, dts, valid: true };
    }

    return { favorites: [], dts: 0, valid: false };
  } catch {
    return { favorites: [], dts: 0, valid: false };
  }
}

function parseRecentsCell(raw) {
  // Accepts:
  //   Legacy bare array of plain IDs:   ["id1","id2",...]
  //   New bare array of RecentItem objects: [{id,workbookId,name},...]
  try {
    if (typeof raw !== "string") raw = String(raw ?? "");
    const s = raw.trim();
    if (!s) return [];
    const v = JSON.parse(s);
    if (!Array.isArray(v)) return [];
    return v.map(e => normalizeRecentItem(e)).filter(Boolean);
  } catch {
    return [];
  }
}

function parseSettingsCell(raw) {
  // Accept either legacy object format: { ...settings }
  // or wrapped format: { dts: <ms>, settings: { ... } }
  try {
    if (typeof raw !== "string") raw = String(raw ?? "");
    const s = raw.trim();
    if (!s) return { settings: {}, dts: 0, valid: false };
    const v = JSON.parse(s);
    if (v && typeof v === "object" && !Array.isArray(v) && v.settings && typeof v.settings === "object" && !Array.isArray(v.settings)) {
      const dts = (typeof v.dts === "number" && isFinite(v.dts)) ? v.dts : 0;
      return { settings: v.settings, dts, valid: true };
    }
    if (v && typeof v === "object" && !Array.isArray(v)) {
      // legacy settings object
      return { settings: v, dts: 0, valid: true };
    }
    return { settings: {}, dts: 0, valid: false };
  } catch {
    return { settings: {}, dts: 0, valid: false };
  }
}

function isValidRuntimePayload(rt) {
  if (!rt || typeof rt !== "object") return false;
  if (typeof rt.dts !== "number" || !isFinite(rt.dts)) return false;
  if (!Array.isArray(rt.favorites)) return false;
  if (!rt.settings || typeof rt.settings !== "object" || Array.isArray(rt.settings)) return false;
  return true;
}

function choosePayload(wbPayload, rtPayload) {
  // Implements LPD rule:
  // - If one invalid, take the other
  // - If both valid, latest dts wins
  // - If tie or within 4 seconds, prefer workbook
  const wbValid = !!wbPayload?.valid;
  const rtValid = !!rtPayload?.valid;

  if (!wbValid && rtValid) return { source: "runtime", payload: rtPayload };
  if (wbValid && !rtValid) return { source: "workbook", payload: wbPayload };
  if (!wbValid && !rtValid) return { source: "workbook", payload: wbPayload };

  const wbDts = typeof wbPayload.dts === "number" ? wbPayload.dts : 0;
  const rtDts = typeof rtPayload.dts === "number" ? rtPayload.dts : 0;
  const diff = Math.abs(wbDts - rtDts);
  if (diff <= 4000) return { source: "workbook", payload: wbPayload };
  return (rtDts > wbDts) ? { source: "runtime", payload: rtPayload } : { source: "workbook", payload: wbPayload };
}

function isPlainObjectEmpty(o) {
  if (!o || typeof o !== "object" || Array.isArray(o)) return true;
  return Object.keys(o).length === 0;
}


async function getOrCreateUserKey() {
  // Prefer OfficeRuntime.storage, fall back to roamingSettings, then localStorage.
  // When found in a lower-priority store, backfill into ORTS so future sessions
  // find it there directly.
  let existing = null;
  const ortsAvailable = typeof OfficeRuntime !== "undefined" && !!OfficeRuntime.storage?.getItem;

  // 1) OfficeRuntime.storage (Shared Runtime)
  try {
    if (ortsAvailable) {
      existing = await OfficeRuntime.storage.getItem(USERKEY_STORAGE_KEY);
      if (existing) {
        identityLog("loaded from OfficeRuntime.storage", existing);
        return existing;
      }
    }
  } catch {}

  // 2) Roaming settings (per-user, persists across sessions)
  try {
    const rs = Office?.context?.roamingSettings;
    if (rs?.get) {
      existing = rs.get(USERKEY_STORAGE_KEY);
      if (existing) {
        identityLog("loaded from roamingSettings", existing);
        // Backfill into ORTS
        try {
          if (ortsAvailable) await OfficeRuntime.storage.setItem(USERKEY_STORAGE_KEY, existing);
        } catch {}
        return existing;
      }
    }
  } catch {}

  // 3) localStorage (last resort; persists per-browser origin)
  try {
    existing = globalThis?.localStorage?.getItem?.(USERKEY_STORAGE_KEY);
    if (existing) {
      identityLog("loaded from localStorage", existing);
      // Backfill into ORTS so future sessions find it there directly
      try {
        if (ortsAvailable) await OfficeRuntime.storage.setItem(USERKEY_STORAGE_KEY, existing);
      } catch {}
      return existing;
    }
  } catch {}

  // Create new key
  const key =
    (globalThis.crypto?.randomUUID?.() ||
      `u_${Date.now()}_${Math.random().toString(16).slice(2)}`);

  identityLog("CREATED NEW userKey", key);

  // Persist to all available backends
  try {
    if (ortsAvailable) await OfficeRuntime.storage.setItem(USERKEY_STORAGE_KEY, key);
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

// --- Read-only detection ---
//
// Combined approach:
//   1. Fast synchronous check via Office.context.document.mode.
//      Office.DocumentMode.ReadOnly = "readOnly" (string, not integer).
//   2. Probe write: attempt to write a value to an existing named item or
//      a scratch cell, catch the error. This catches Protected View, co-auth
//      without edit rights, and any case document.mode misses.
//
// detectWorkbookReadOnly() returns true if the workbook is not writable.
// Call once at startup; cache result in commands.js.

export async function detectWorkbookReadOnly() {
  // Primary: load the workbook's readOnly property directly from Excel.
  // This is the Office JS equivalent of VBA's Workbook.ReadOnly.
  try {
    const result = await Excel.run(async (context) => {
      context.workbook.load("readOnly");
      await context.sync();
      return context.workbook.readOnly;
    });
    if (result === true) return true;
  } catch {
    // ignore — fall through to document.mode check
  }

  // Fallback: Office document mode string check.
  try {
    const mode = Office?.context?.document?.mode;
    if (mode === "readOnly" || mode === Office?.DocumentMode?.ReadOnly) return true;
  } catch {
    // ignore
  }

  return false;
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

// Read-only variant: returns the settings sheet if it exists, null if it doesn't.
// Never creates the sheet, never modifies visibility.
async function getSettingsSheetIfExists(context) {
  const ws = context.workbook.worksheets.getItemOrNullObject(SETTINGS_SHEET_NAME);
  ws.load("name,visibility");
  await context.sync();
  if (ws.isNullObject) return null;
  return ws;
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

// Read-only variant: reads identity from the sheet without writing anything.
// Returns { workbookGuid, filenameFingerprint } with nulls if missing/invalid.
async function readWorkbookIdentity(context, settingsSheet) {
  const range = settingsSheet.getRange(WB_ID_RANGE_ADDRESS);
  range.load("values");
  await context.sync();

  const values = range.values || [["", ""], ["", ""]];
  const b1 = (values[0]?.[1] || "").toString();
  const b2 = (values[1]?.[1] || "").toString();

  const guid = isValidGuid(b1) ? b1.trim() : null;
  const filenameFingerprint = b2 || null;

  return { workbookGuid: guid, filenameFingerprint };
}

export async function getWorkbookIdentity() {
  return Excel.run(async (context) => {
    const settingsSheet = await ensureSettingsSheet(context);
    return ensureWorkbookIdentity(context, settingsSheet);
  });
}

export async function getWorkbookIdentityReadOnly() {
  return Excel.run(async (context) => {
    const settingsSheet = await getSettingsSheetIfExists(context);
    if (!settingsSheet) return { workbookGuid: null, filenameFingerprint: null };
    return readWorkbookIdentity(context, settingsSheet);
  });
}

async function getUserColumn(context, settingsSheet, userKey) {
  // Search row 1 starting from column D for userKey; else first empty.
  // We'll scan G1:ZZ1; cols A-F are reserved for inventory identity and future use.
  const headerRange = settingsSheet.getRange("G1:ZZ1");
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
  const colIdx1 = 7 + offset;
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

  const favRaw = favCell.values?.[0]?.[0];
  const recRaw = recCell.values?.[0]?.[0];
  const setRaw = setCell.values?.[0]?.[0];

  const favParsed = parseFavoritesCell(favRaw);
  const setParsed = parseSettingsCell(setRaw);
  const recents = parseRecentsCell(recRaw);
  const recentsValid = Array.isArray(recents);

  // We treat Favorites + Settings as a single logical payload for reconciliation.
  // Use max dts to be robust if they ever diverge.
  const dts = Math.max(Number(favParsed.dts || 0), Number(setParsed.dts || 0));

  return {
    favorites: favParsed.favorites,
    recents,
    settings: setParsed.settings,
    __meta: {
      favoritesValid: favParsed.valid,
      settingsValid: setParsed.valid,
      recentsValid,
      dts
    }
  };
}



async function writeUserCells(context, sheet, colLetter, { favorites, recents, settings, dtsOverride }) {
  const favCell = sheet.getRange(`${colLetter}${ROW_FAVORITES}`);
  const recCell = sheet.getRange(`${colLetter}${ROW_RECENTS}`);
  const setCell = sheet.getRange(`${colLetter}${ROW_SETTINGS}`);

  const dts = (typeof dtsOverride === "number" && isFinite(dtsOverride)) ? dtsOverride : Date.now();

  // favorites is always an array of FavoriteItem objects
  const favItems = Array.isArray(favorites) ? favorites : [];
  const favPayload = { dts, favorites: favItems };
  const setPayload = { dts, settings: (settings && typeof settings === "object" && !Array.isArray(settings)) ? settings : {} };

  if (favorites !== undefined) {
    favCell.values = [[safeJsonStringify(favPayload)]];
  }
  if (recents !== undefined) {
    // recents is always an array of RecentItem objects — no wrapper, matches XLL format
    recCell.values = [[safeJsonStringify(Array.isArray(recents) ? recents : [])]];
  }
  if (settings !== undefined) {
    setCell.values = [[safeJsonStringify(setPayload)]];
  }
  await context.sync();
}

async function loadInventory(context, sheet, userColLetter) {
  // Load A52:C2000 and user's frequency column range for same rows.
  // Column A = Office.js sheet.id, B = sheet.name, C = VBA CodeName (COM). User col stores frequency.
  // Also load the dev premium flag cell (A10) in the same batch — zero added latency.
  const endRow = 2000;
  const invRange = sheet.getRange(`A${INV_START_ROW}:C${endRow}`);
  const freqRange = sheet.getRange(`${userColLetter}${INV_START_ROW}:${userColLetter}${endRow}`);
  const devFlagCell = sheet.getRange(DEV_FLAG_CELL);
  invRange.load("values");
  freqRange.load("values");
  devFlagCell.load("values");
  await context.sync();

  const devPremium = String(devFlagCell.values?.[0]?.[0] ?? "").trim() === DEV_FLAG_VALUE;

  const inv = invRange.values || [];
  const freq = freqRange.values || [];
  const nowMs = Date.now();
  const rows = [];
  for (let i = 0; i < inv.length; i++) {
    const rowNum = INV_START_ROW + i;
    const id = inv[i]?.[0] ?? "";
    const name = inv[i]?.[1] ?? "";
    const raw = freq[i]?.[0];
    const { freq: storedFreq, dts } = parseFreqCell(raw);
    const decayedFreq = decayFreq(storedFreq, dts, nowMs);
    const codename = inv[i]?.[2] ?? "";
    rows.push({ rowNum, id: String(id || ""), name: String(name || ""), codename: String(codename || ""), freq: decayedFreq, storedFreq, dts });
  }
  return { rows, devPremium };
}

async function syncInventoryWithVisibleSheets(context, sheet, userColLetter, visibleSheets, allSheets) {
  // visibleSheets: [{id,name,orderIndex?}] — sheets currently visible (used to add/update rows)
  // allSheets: [{id,name}] — ALL sheets in workbook including hidden (used to decide what to clear)
  // Rows are only cleared for sheets that are completely absent from the workbook.
  // Hidden sheets retain their inventory row and frequency data intact.
  const { rows } = await loadInventory(context, sheet, userColLetter);

  // Build set of all workbook sheet ids and names (visible + hidden) for existence check
  const allSheetIds = new Set((allSheets || visibleSheets).map(s => String(s.id || "")));
  const allSheetNames = new Set((allSheets || visibleSheets).map(s => String(s.name || "")));

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
    // initialize freq to { freq: 0, dts: 0 }
    const fCell = sheet.getRange(`${userColLetter}${lastUsedRow}`);
    fCell.values = [[stringifyFreqCell(0, 0)]];
    matchedRows.add(lastUsedRow);
  }

  // Clear rows only for sheets that no longer exist anywhere in the workbook (not merely hidden).
  for (const r of rows) {
    if (!(r.id || r.name)) continue;
    if (matchedRows.has(r.rowNum)) continue;
    // Only clear if the sheet is gone from the workbook entirely
    const existsById = r.id && allSheetIds.has(r.id);
    const existsByName = r.name && allSheetNames.has(r.name);
    if (!existsById && !existsByName) {
      sheet.getRange(`A${r.rowNum}:C${r.rowNum}`).clear();
      sheet.getRange(`${userColLetter}${r.rowNum}`).clear();
    }
  }

  await context.sync();
}

async function incrementFrequency(context, sheet, userColLetter, sheetId) {
  const { rows } = await loadInventory(context, sheet, userColLetter);
  const target = rows.find(r => r.id === sheetId) || null;
  if (!target) return 0;

  const cell = sheet.getRange(`${userColLetter}${target.rowNum}`);
  cell.load("values");
  await context.sync();

  // Decay the existing stored value forward to now, then add 1.
  // This keeps the stored freq meaningful as "decayed count as of dts".
  const nowMs = Date.now();
  const { freq: storedFreq, dts } = parseFreqCell(cell.values?.[0]?.[0]);
  const decayed = decayFreq(storedFreq, dts, nowMs);
  const next = decayed + 1;
  cell.values = [[stringifyFreqCell(next, nowMs)]];
  await context.sync();
  return next;
}

export async function getJumpToState(options = {}) {
  const preferCache = !!options?.preferCache;
  const isReadOnly = !!options?.isReadOnly;
  const userKey = await getOrCreateUserKey();
  if (!userKey) {
    return { userKey: null, sheets: [], favorites: [], recents: [], settings: {}, global: {}, isReadOnly };
  }

  // --- Read-only path ---
  // Never writes to the workbook. Attempts to read from the settings sheet if it exists,
  // falls back to RT caches (RT10 for favorites/settings, RT3 for sheets/recents/freq).
  if (isReadOnly) {
    let wbId = { workbookGuid: null, filenameFingerprint: null };
    let favorites = [];
    let recents = [];
    let settings = {};
    let sheets = [];
    let freqById = {};
    let devPremium = false;
    let __meta = { favoritesValid: false, settingsValid: false, recentsValid: false, dts: 0 };

    // Try to read from the settings sheet if it exists
    try {
      const wbData = await Excel.run(async (context) => {
        const settingsSheet = await getSettingsSheetIfExists(context);

        // Always read visible sheets — regardless of whether the settings sheet exists.
        const wsItems = context.workbook.worksheets;
        wsItems.load("items/id,name,visibility");
        await context.sync();
        const visible = wsItems.items.filter(ws => ws.visibility === Excel.SheetVisibility.visible);
        const visibleSheets = visible.map((ws, idx) => ({ id: ws.id, name: ws.name, orderIndex: idx }));

        if (!settingsSheet) return { identity: { workbookGuid: null, filenameFingerprint: null }, visibleSheets, userCells: null, invRows: [], devPremium: false };

        const identity = await readWorkbookIdentity(context, settingsSheet);

        // Only attempt to read user cells if we have a valid GUID (so we can find the user column)
        let userCells = null;
        let invRows = [];
        let dp = false;
        if (identity.workbookGuid) {
          try {
            // getUserColumn in read-only: find existing column but don't create one if absent
            const headerRange = settingsSheet.getRange("G1:ZZ1");
            headerRange.load("values");
            await context.sync();
            const headerVals = headerRange.values?.[0] || [];
            const foundOffset = headerVals.findIndex(v => v === userKey);
            if (foundOffset >= 0) {
              const colIdx1 = 7 + foundOffset;
              const colLetter = colIndexToLetter(colIdx1);
              userCells = await readUserCells(context, settingsSheet, colLetter);
              const inv = await loadInventory(context, settingsSheet, colLetter);
              invRows = inv.rows;
              dp = inv.devPremium;
            }
          } catch {
            // ignore — best-effort read
          }
        }

        return { identity, visibleSheets, userCells, invRows, devPremium: dp };
      });

      if (wbData) {
        wbId = wbData.identity;
        sheets = wbData.visibleSheets || [];
        devPremium = wbData.devPremium;

        if (wbData.userCells) {
          favorites = wbData.userCells.favorites || [];
          recents = wbData.userCells.recents || [];
          settings = wbData.userCells.settings || {};
          __meta = wbData.userCells.__meta || __meta;
        }

        for (const r of (wbData.invRows || [])) {
          if (r.id) freqById[r.id] = Number(r.freq || 0);
        }
      }
    } catch {
      // Settings sheet unreadable — fall through to cache-only
    }

    // Supplement with RT caches if workbook read gave us nothing useful
    const { workbookGuid, filenameFingerprint } = wbId;
    if (workbookGuid && filenameFingerprint) {
      // RT10: favorites + settings safety net
      if (!__meta.favoritesValid || !__meta.settingsValid) {
        try {
          const rt = await rt10Read(workbookGuid, filenameFingerprint);
          if (isValidRuntimePayload(rt)) {
            if (!__meta.favoritesValid && Array.isArray(rt.favorites)) favorites = rt.favorites.map((e, i) => normalizeFavoriteItem(e, i)).filter(Boolean);
            if (!__meta.settingsValid && rt.settings) settings = rt.settings;
          }
        } catch {}
      }

      // RT3: sheets + recents + freq (cache-only, allowed to be stale)
      if (sheets.length === 0) {
        try {
          const perf = await rt3Read(workbookGuid, filenameFingerprint);
          if (perf && Array.isArray(perf.sheets)) {
            sheets = perf.sheets;
            if (!__meta.recentsValid && Array.isArray(perf.recents)) recents = perf.recents.map(r => normalizeRecentItem(r)).filter(Boolean);
            if (perf.freqById && typeof perf.freqById === "object") freqById = perf.freqById;
            if (perf.devPremium) devPremium = true;
          }
        } catch {}
      }
    }

    const idToName = new Map(sheets.map(s => [s.id, s.name]));
    const favIds = Array.isArray(favorites) ? favorites : [];
    const recIds = Array.isArray(recents) ? recents : [];

    return {
      __wbId: wbId,
      userKey,
      sheets,
      favorites: favIds.map(f => ({ ...f, name: idToName.get(f.id) || f.name || "" })),
      recents: recIds.map(r => ({ ...r, name: idToName.get(r.id) || r.name || "" })),
      settings: (settings && typeof settings === "object") ? settings : {},
      __meta,
      global: { freqById, devPremium },
      isReadOnly: true,
    };
  }

  // --- Normal (writable) path ---

  // Phase 4 (perf): If requested, try a fast path that avoids enumerating worksheets.
  // We still read the per-user cells from the workbook (small), then use RT3 cache
  // for the visible worksheet list + name mapping (allowed to be stale).
  if (preferCache) {
    const mini = await Excel.run(async (context) => {
      const settingsSheet = await ensureSettingsSheet(context);
      const wbId = await ensureWorkbookIdentity(context, settingsSheet);
      const { colLetter } = await getUserColumn(context, settingsSheet, userKey);
      const { favorites, recents, settings, __meta } = await readUserCells(context, settingsSheet, colLetter);

      return {
        __wbId: wbId,
        userKey,
        favorites,
        recents,
        settings: (settings && typeof settings === "object") ? settings : {},
        __meta,
        global: { freqById: {} }
      };
    });

    const { workbookGuid, filenameFingerprint } = mini?.__wbId || {};
    if (workbookGuid && filenameFingerprint) {
      const perf = await rt3Read(workbookGuid, filenameFingerprint);
      if (perf && Array.isArray(perf.sheets)) {
        const sheets = perf.sheets;
        const idToName = new Map(sheets.map((s) => [s.id, s.name]));

        const favIds = Array.isArray(mini.favorites) ? mini.favorites : [];
        const recIds = Array.isArray(mini.recents) ? mini.recents : [];

        return {
          __wbId: mini.__wbId,
          userKey: mini.userKey,
          sheets,
          favorites: favIds.map((f) => ({ ...f, name: idToName.get(f.id) || f.name || "" })),
          recents: recIds.map((r) => ({ ...r, name: idToName.get(r.id) || r.name || "" })),
          settings: mini.settings,
          __meta: mini.__meta,
          global: {
            freqById: (perf.freqById && typeof perf.freqById === "object") ? perf.freqById : {},
            devPremium: !!(perf.devPremium)
          }
        };
      }
    }
    // If no perf cache is available, fall through to full workbook-backed state.
  }

  // Full workbook state (authoritative for visible sheets and inventory sync).
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
    const allWorkbookSheets = sheets.items.map(ws => ({ id: ws.id, name: ws.name }));

    // Reconcile inventory and read per-user blobs
    await syncInventoryWithVisibleSheets(context, settingsSheet, colLetter, visibleSheets, allWorkbookSheets);
    const { favorites, recents, settings, __meta } = await readUserCells(context, settingsSheet, colLetter);

    // Build enriched favorites/recents objects with live names
    const idToName = new Map(visibleSheets.map(s => [s.id, s.name]));
    const favObjs = (Array.isArray(favorites) ? favorites : []).map(f => ({
      ...f,
      name: idToName.get(f.id) || f.name || ""
    }));
    const recObjs = (Array.isArray(recents) ? recents : []).map(r => ({
      ...r,
      name: idToName.get(r.id) || r.name || ""
    }));

    // Load frequency values and dev flag for visible sheets (single batch, no extra sync).
    const { rows: invRows, devPremium } = await loadInventory(context, settingsSheet, colLetter);
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
      __meta,
      global: { freqById, devPremium }
    };
  });

  // Phase 3: reconcile workbook vs runtime (Favorites + Settings only) using dts rules, then self-heal in background.
  const { workbookGuid, filenameFingerprint } = wb?.__wbId || {};
  let final = wb;

  if (workbookGuid && filenameFingerprint) {
    const rt = await rt10Read(workbookGuid, filenameFingerprint);

    const wbFavValid = !!wb?.__meta?.favoritesValid;
    const wbSetValid = !!wb?.__meta?.settingsValid;
    const wbDts = Number(wb?.__meta?.dts || 0);

    // Runtime payload is the safety-net for Favorites + Settings.
    // Validate using the actual runtime schema (it does not carry workbook-style __meta flags).
    const rtValid = isValidRuntimePayload(rt);
    const rtDts = Number(rt?.dts || 0);

    const wbValid = wbFavValid && wbSetValid;

    const choose = (() => {
      if (!wbValid && rtValid) return "rt";
      if (wbValid && !rtValid) return "wb";
      if (!wbValid && !rtValid) return "wb"; // both bad: prefer workbook to avoid silent surprises
      // both valid:
      if (Math.abs(rtDts - wbDts) <= 4000) return "wb";
      if (rtDts > wbDts) return "rt";
      return "wb";
    })();

    if (choose === "rt" && rt) {
      // Apply runtime state to the returned object
      const rtFavItems = Array.isArray(rt.favorites) ? rt.favorites : [];
      const rtSet = (rt.settings && typeof rt.settings === "object") ? rt.settings : {};
      const idToName = new Map((Array.isArray(wb?.sheets) ? wb.sheets : []).map((s) => [s.id, s.name]));
      final = {
        ...wb,
        favorites: rtFavItems.map((f) => ({ ...f, name: idToName.get(f.id) || f.name || "" })),
        settings: rtSet,
        __meta: { ...(wb.__meta || {}), dts: rtDts, favoritesValid: true, settingsValid: true }
      };

      // Background self-heal: write runtime-chosen state back into workbook (best effort)
      setTimeout(() => {
        Excel.run(async (context) => {
          const settingsSheet = await ensureSettingsSheet(context);
          const { colLetter } = await getUserColumn(context, settingsSheet, userKey);
          const current = await readUserCells(context, settingsSheet, colLetter);
          await writeUserCells(context, settingsSheet, colLetter, {
            favorites: rtFavItems,
            recents: current.recents,
            settings: rtSet,
            dtsOverride: rtDts
          });
        }).catch(() => {});
      }, 0);
    }

    if (choose === "wb" && wbValid) {
      // Background self-heal: write workbook-chosen state into runtime (best effort)
      const wbFavItems = Array.isArray(wb?.favorites) ? wb.favorites : [];
      const wbSet = wb?.settings || {};
      const dts = wbDts || Date.now();
      setTimeout(() => {
        rt10Write(workbookGuid, filenameFingerprint, wbFavItems, wbSet, dts).catch(() => {});
      }, 0);
    }
  }

  // Phase 4: write perf cache (RT3) for visible worksheets + recents (cache-only).
  if (workbookGuid && filenameFingerprint) {
    try {
      const sheets = Array.isArray(final?.sheets) ? final.sheets : [];
      const recItems = Array.isArray(final?.recents) ? final.recents : [];
      await rt3Write(workbookGuid, filenameFingerprint, {
        dts: Date.now(),
        sheets,
        recents: recItems,
        freqById: final?.global?.freqById || {},
        devPremium: !!(final?.global?.devPremium)
      });
    } catch {
      // ignore
    }
  }

  return { ...final, isReadOnly: false };
}


export async function toggleFavorite(sheetId) {
  const userKey = await getOrCreateUserKey();
  if (!userKey) return null;

  return Excel.run(async (context) => {
    const sheet = await ensureSettingsSheet(context);
    const wbId = await ensureWorkbookIdentity(context, sheet);
    const { colLetter } = await getUserColumn(context, sheet, userKey);
    const state = await readUserCells(context, sheet, colLetter);
    const favs = Array.isArray(state.favorites) ? [...state.favorites] : [];

    const idx = favs.findIndex(f => f.id === sheetId);
    if (idx >= 0) {
      favs.splice(idx, 1);
      // Re-derive digit and displayOrder from new positions
      favs.forEach((f, i) => { f.digit = i + 1; f.displayOrder = i + 1; });
    } else {
      const newPos = favs.length;
      favs.push({ id: sheetId, workbookId: "this", digit: newPos + 1, displayOrder: newPos + 1, name: "" });
      if (favs.length > MAX_FAVORITES) favs.length = MAX_FAVORITES;
    }
    const dts = Date.now();
    const wbSettings = pickWbSettings(state.settings);
    await writeUserCells(context, sheet, colLetter, { favorites: favs, recents: state.recents, settings: wbSettings, dtsOverride: dts });

    // Mirror into runtime safety net (best effort)
    const { workbookGuid, filenameFingerprint } = wbId || {};
    if (workbookGuid && filenameFingerprint) {
      await rt10Write(workbookGuid, filenameFingerprint, favs, wbSettings, dts);
    }

    return favs;
  });
}


export async function setFavorites(favItems) {
  const userKey = await getOrCreateUserKey();
  if (!userKey) return;

  const nextFavs = Array.isArray(favItems) ? favItems.filter(f => f && f.id) : [];
  const dts = Date.now();

  const out = await Excel.run(async (context) => {
    const settingsSheet = await ensureSettingsSheet(context);
    const wbId = await ensureWorkbookIdentity(context, settingsSheet);
    const { colLetter } = await getUserColumn(context, settingsSheet, userKey);
    const { recents, settings, __meta } = await readUserCells(context, settingsSheet, colLetter);
    const setValid = !!__meta?.settingsValid;
    await writeUserCells(context, settingsSheet, colLetter, { favorites: nextFavs, recents, settings: setValid ? pickWbSettings(settings) : undefined, dtsOverride: dts });
    return { wbId, settings, __meta };
  });

  const { workbookGuid, filenameFingerprint } = out?.wbId || {};
  if (workbookGuid && filenameFingerprint) {
    const wbSetValid = !!out?.__meta?.settingsValid;
    let nextSettings = (out?.settings && typeof out.settings === "object") ? out.settings : {};
    if (!wbSetValid) {
      const rt = await rt10Read(workbookGuid, filenameFingerprint);
      if (isValidRuntimePayload(rt)) {
        nextSettings = (rt.settings && typeof rt.settings === "object") ? rt.settings : {};
      }
    }
    await rt10Write(workbookGuid, filenameFingerprint, nextFavs, nextSettings, dts);
  }
}

export async function addFavorite(sheetId) {
  if (!sheetId) return null;
  const state = await getJumpToState();
  const current = Array.isArray(state.favorites) ? state.favorites : [];
  if (current.some(f => f.id === sheetId)) return current;
  const newPos = current.length;
  const next = [...current, { id: sheetId, workbookId: "this", digit: newPos + 1, displayOrder: newPos + 1, name: "" }]
    .slice(0, MAX_FAVORITES);
  return setFavorites(next);
}

export async function removeFavorite(sheetId) {
  if (!sheetId) return null;
  const state = await getJumpToState();
  const current = Array.isArray(state.favorites) ? state.favorites : [];
  const next = current.filter(f => f.id !== sheetId);
  // Re-derive digit and displayOrder from new positions
  next.forEach((f, i) => { f.digit = i + 1; f.displayOrder = i + 1; });
  return setFavorites(next);
}

export async function moveFavorite(sheetId, direction) {
  if (!sheetId) return null;
  if (direction !== "up" && direction !== "down") return null;

  const state = await getJumpToState();
  const current = Array.isArray(state.favorites) ? state.favorites : [];
  const idx = current.findIndex(f => f.id === sheetId);
  if (idx < 0) return current;

  const to = direction === "up" ? idx - 1 : idx + 1;
  if (to < 0 || to >= current.length) return current;

  const next = current.slice();
  const [item] = next.splice(idx, 1);
  next.splice(to, 0, item);
  // Re-derive digit and displayOrder from new positions
  next.forEach((f, i) => { f.digit = i + 1; f.displayOrder = i + 1; });
  return setFavorites(next);
}




export async function setUiSettings(nextSettings) {
  const userKey = await getOrCreateUserKey();
  if (!userKey) return;

  const incoming = (nextSettings && typeof nextSettings === "object" && !Array.isArray(nextSettings)) ? nextSettings : {};
  // Only persist workbook-scoped settings in the workbook Settings sheet.
  const normalized = {};
  for (const k of WB_SETTINGS_KEYS) {
    if (Object.prototype.hasOwnProperty.call(incoming, k)) normalized[k] = incoming[k];
  }
  const dts = Date.now();

  const out = await Excel.run(async (context) => {
    const settingsSheet = await ensureSettingsSheet(context);
    const wbId = await ensureWorkbookIdentity(context, settingsSheet);
    const { colLetter } = await getUserColumn(context, settingsSheet, userKey);
    const { favorites, recents, settings: existingSettings, __meta } = await readUserCells(context, settingsSheet, colLetter);
    const favValid = !!__meta?.favoritesValid;
    const prev = (existingSettings && typeof existingSettings === "object" && !Array.isArray(existingSettings)) ? existingSettings : {};
    const prevFiltered = {};
    for (const k of WB_SETTINGS_KEYS) {
      if (Object.prototype.hasOwnProperty.call(prev, k)) prevFiltered[k] = prev[k];
    }
    const mergedSettings = { ...prevFiltered, ...normalized };
    await writeUserCells(context, settingsSheet, colLetter, { favorites: favValid ? favorites : undefined, recents, settings: mergedSettings, dtsOverride: dts });
    return { wbId, favorites, __meta };
  });

  const { workbookGuid, filenameFingerprint } = out?.wbId || {};
  if (workbookGuid && filenameFingerprint) {
    const wbFavValid = !!out?.__meta?.favoritesValid;
    let favIds = Array.isArray(out?.favorites) ? out.favorites.filter(Boolean) : [];

    // If workbook favorites were missing/invalid, preserve runtime favorites instead of wiping.
    if (!wbFavValid) {
      const rt = await rt10Read(workbookGuid, filenameFingerprint);
      if (isValidRuntimePayload(rt)) {
        favIds = Array.isArray(rt.favorites) ? rt.favorites.filter(Boolean) : [];
      }
    }

    await rt10Write(workbookGuid, filenameFingerprint, favIds, normalized, dts);
  }
}

export async function recordActivation(sheetId) {
  const userKey = await getOrCreateUserKey();
  if (!userKey) return null;

  const out = await Excel.run(async (context) => {
    const settingsSheet = await ensureSettingsSheet(context);
    const wbId = await ensureWorkbookIdentity(context, settingsSheet);
    const { colLetter } = await getUserColumn(context, settingsSheet, userKey);

    // Get current visible sheets (for reconciliation)
    const ws = context.workbook.worksheets;
    ws.load("items/id,name,visibility");
    await context.sync();
    const visible = ws.items.filter(w => w.visibility === Excel.SheetVisibility.visible);
    const visibleSheets = visible.map((w, idx) => ({ id: w.id, name: w.name, orderIndex: idx }));
    const allWorkbookSheets = ws.items.map(w => ({ id: w.id, name: w.name }));

    await syncInventoryWithVisibleSheets(context, settingsSheet, colLetter, visibleSheets, allWorkbookSheets);

    const state = await readUserCells(context, settingsSheet, colLetter);

    // Update recents — stored as RecentItem objects
    const rec = Array.isArray(state.recents) ? [...state.recents] : [];
    const existing = rec.findIndex(r => r.id === sheetId);
    if (existing >= 0) rec.splice(existing, 1);
    rec.unshift({ id: sheetId, workbookId: "this", name: "" });
    if (rec.length > MAX_RECENTS) rec.length = MAX_RECENTS;

    const favValid = !!state?.__meta?.favoritesValid;
    const setValid = !!state?.__meta?.settingsValid;
    await writeUserCells(context, settingsSheet, colLetter, {
      favorites: favValid ? state.favorites : undefined,
      settings: setValid ? state.settings : undefined,
      recents: rec
    });
    const nextFreq = await incrementFrequency(context, settingsSheet, colLetter, sheetId);

    return { wbId, recents: rec, freq: nextFreq };
  });

  // Phase 4: mirror recents into RT3 cache (best effort, cache-only).
  const { workbookGuid, filenameFingerprint } = out?.wbId || {};
  if (workbookGuid && filenameFingerprint) {
    try {
      const perf = await rt3Read(workbookGuid, filenameFingerprint) || {};
      const next = {
        ...perf,
        dts: Date.now(),
        recents: Array.isArray(out.recents) ? out.recents : perf.recents
      };
      await rt3Write(workbookGuid, filenameFingerprint, next);
    } catch {
      // ignore
    }
  }

  return { recents: out?.recents, freq: out?.freq };
}
