import * as React from "react";
import PropTypes from "prop-types";
import {
  makeStyles,
  Input,
  Text,
  Button,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  TabList,
  Tab,
  Switch,
  Checkbox,
} from "@fluentui/react-components";
import {
  MoreHorizontal24Regular,
  Dismiss24Regular,
  ReOrderDotsVertical24Regular,
} from "@fluentui/react-icons";

import { MAX_RECENTS } from "../../shared/constants";

/* =========================
   Constants
========================= */

console.log("### JUMPTO APP VERSION: 2026-01-08 A ###");

const SETTINGS_SHEET_NAME = "_JumpToAddinSettings";
const USER_KEY_STORAGE_KEY = "JumpToAddin.UserKey";
const USER_COL_START = 4; // column D = 4 (1-based)
const USER_KEY_ROW = 1;
const USER_BLOB_ROW = 2;

const MAX_FAVORITES = 20;

// Inventory table starts at row 52 (1-based) => 51 (0-based)
const INVENTORY_START_ROW0 = 51; // row 52
const INVENTORY_NAME_COL0 = 0; // column A
const INVENTORY_META_COLS = 3; // columns A-C reserved (A=name, B/C reserved)

// Favorites hotkey labels: 1..9,0 then blank
const FAVORITE_HOTKEY_LABELS = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "0"];

// Per-user, per-workbook option (stored in user blob)
const DEFAULT_ONE_DIGIT_ACTIVATION_ENABLED = true;

/* =========================
   Styles
========================= */

const useStyles = makeStyles({
  root: { minHeight: "100vh" },
  section: { padding: "12px 16px" },

  list: { marginTop: "10px", borderTop: "1px solid rgba(0,0,0,0.08)" },

  row: {
    padding: "8px 0",
    borderBottom: "1px solid rgba(0,0,0,0.06)",
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    gap: "10px",
  },

  rowLeft: { display: "flex", alignItems: "center", gap: "10px", flex: 1, minWidth: 0 },
  rowRight: { display: "flex", alignItems: "center", gap: "8px" },

  hotkey: {
    width: "18px",
    textAlign: "right",
    opacity: 0.7,
    fontVariantNumeric: "tabular-nums",
    userSelect: "none",
    flex: "0 0 auto",
  },

  sheetName: {
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },

  meta: { opacity: 0.75, fontSize: "12px", marginTop: "6px" },
  subheader: { marginTop: "14px" },

  manageHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "12px",
  },

  favoritesDnDRow: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    padding: "8px 10px",
    borderRadius: "8px",
    border: "1px solid rgba(0,0,0,0.08)",
    marginTop: "8px",
    background: "rgba(255,255,255,0.6)",
  },
  dragHandle: { opacity: 0.7, cursor: "grab", flex: "0 0 auto" },
  dndName: { flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },

  pickerBox: {
    marginTop: "10px",
    borderRadius: "10px",
    border: "1px solid rgba(0,0,0,0.08)",
    padding: "10px",
    background: "rgba(255,255,255,0.65)",
  },
  pickerList: {
    marginTop: "8px",
    maxHeight: "240px",
    overflowY: "auto",
    borderTop: "1px solid rgba(0,0,0,0.06)",
  },
  pickerRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "12px",
    padding: "8px 2px",
    borderBottom: "1px solid rgba(0,0,0,0.05)",
  },
  pickerName: { flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
});

/* =========================
   Utilities
========================= */

function normalizeBlob(obj) {
  const safe = obj && typeof obj === "object" ? obj : {};
  const options = safe.options && typeof safe.options === "object" ? safe.options : {};
  return {
    favorites: Array.isArray(safe.favorites) ? safe.favorites.filter(Boolean) : [],
    recents: Array.isArray(safe.recents) ? safe.recents.filter(Boolean) : [],
    options: {
      oneDigitActivationEnabled:
        typeof options.oneDigitActivationEnabled === "boolean"
          ? options.oneDigitActivationEnabled
          : DEFAULT_ONE_DIGIT_ACTIVATION_ENABLED,
      ...options,
    },
  };
}

function pushUniqueFront(list, item, maxLen) {
  const out = [item, ...list.filter((x) => x !== item)];
  return out.slice(0, maxLen);
}

function toNumberOrZero(v) {
  const n = typeof v === "number" ? v : Number(v);
  return Number.isFinite(n) ? n : 0;
}

function stripLeadingSpaces(s) {
  return String(s ?? "").replace(/^\s+/, "");
}

function digitToFavoriteIndex1Based(digitChar) {
  if (digitChar === "0") return 10;
  const n = Number(digitChar);
  if (Number.isInteger(n) && n >= 1 && n <= 9) return n;
  return null;
}

function moveArrayItem(arr, fromIndex, toIndex) {
  const a = [...arr];
  const [item] = a.splice(fromIndex, 1);
  a.splice(toIndex, 0, item);
  return a;
}

function isMissingSheetError(err) {
  const code = err?.code || err?.name;
  const msg = String(err?.message || "");
  return (
    code === "ItemNotFound" ||
    /ItemNotFound|not found|does not exist|Cannot find|doesn't exist|no longer exists/i.test(msg)
  );
}

async function hideTaskpaneSafely() {
  try {
    // Newer API (some hosts)
    if (globalThis.Office?.addin?.hide) {
      await Office.addin.hide();
      console.log("hideTaskpaneSafely: Office.addin.hide() succeeded");
      return true;
    }
  } catch (e) {
    console.warn("hideTaskpaneSafely: Office.addin.hide() failed", e);
  }

  try {
    // Excel taskpane close API (commonly supported)
    if (globalThis.Office?.context?.ui?.closeContainer) {
      Office.context.ui.closeContainer();
      console.log("hideTaskpaneSafely: Office.context.ui.closeContainer() called");
      return true;
    }
  } catch (e) {
    console.warn("hideTaskpaneSafely: closeContainer failed", e);
  }

  console.warn("hideTaskpaneSafely: no supported close/hide API found");
  return false;
}

// TEMP (Shared Runtime proof): set true to log SharedRuntime support + visibility events.
// Keep this false in production.
const DEBUG_SHARED_RUNTIME_PROBE = false;
let __sharedRuntimeProbeInstalled = false;

async function installSharedRuntimeProbeOnce() {
  if (!DEBUG_SHARED_RUNTIME_PROBE || __sharedRuntimeProbeInstalled) return;
  __sharedRuntimeProbeInstalled = true;

  await Office.onReady();

  const req = Office.context.requirements;
  const diag = Office.context.diagnostics;

  console.group("JTS SharedRuntime probe");
  console.log("host:", Office.context.host);
  console.log("platform:", Office.context.platform);
  console.log("Office build (diagnostics.version):", diag?.version);
  console.log(
    "SharedRuntime 1.1 supported?:",
    req?.isSetSupported?.("SharedRuntime", "1.1")
  );
  console.log("typeof Office.addin:", typeof Office.addin);
  console.log("typeof Office.addin.hide:", typeof Office.addin?.hide);
  console.log(
    "typeof Office.addin.onVisibilityModeChanged:",
    typeof Office.addin?.onVisibilityModeChanged
  );
  console.log("Office.VisibilityMode:", Office.VisibilityMode);
  console.groupEnd();

  // Visibility event probe (best-effort).
  try {
    const unsubscribe = await Office.addin.onVisibilityModeChanged((msg) => {
      console.log("JTS VisibilityModeChanged:", msg?.visibilityMode, msg);
    });
    // Expose unsubscribe for manual cleanup in devtools if you want it:
    window.__JTS_unsubVisibility = unsubscribe;
    console.log("JTS: onVisibilityModeChanged handler attached.");
  } catch (e) {
    console.error("JTS: onVisibilityModeChanged attach failed:", e);
  }
}


/**
 * Frequency significance (tiering) so small differences don't reorder the list.
 * Rules:
 * - freq < 10 => tier 0 (ignore frequency)
 * - else tiers grow by 1.35x increments from 10
 */
function freqTier(freq, minCount = 10, ratio = 1.35) {
  const f = Number(freq) || 0;
  if (f < minCount) return 0;
  // Tier 1 starts at minCount; tier increases when we multiply by ratio
  return 1 + Math.floor(Math.log(f / minCount) / Math.log(ratio));
}

async function officeReady() {
  if (!globalThis.Office) {
    throw new Error("Office.js not available. Are you running inside Excel?");
  }
  await new Promise((resolve) => Office.onReady(() => resolve()));
}

/* =========================
   Blob helpers (favorites/recents/options)
========================= */

async function readUserBlobAtAddress(blobAddress) {
  if (!globalThis.Excel) throw new Error("Excel.js not available. Are you running inside Excel?");
  await officeReady();

  return Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(SETTINGS_SHEET_NAME);
    const cell = ws.getRange(blobAddress);
    cell.load("values");
    await context.sync();

    const raw = cell.values?.[0]?.[0];
    const text = typeof raw === "string" ? raw.trim() : "";

    if (!text) return normalizeBlob(null);

    try {
      return normalizeBlob(JSON.parse(text));
    } catch {
      return normalizeBlob(null);
    }
  });
}

async function updateUserBlobAtAddress(blobAddress, updater) {
  if (!blobAddress) throw new Error("Blob address not initialized yet.");
  if (!globalThis.Excel) throw new Error("Excel.js not available. Are you running inside Excel?");
  await officeReady();

  return Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(SETTINGS_SHEET_NAME);
    const cell = ws.getRange(blobAddress);

    cell.load("values");
    await context.sync();

    const raw = cell.values?.[0]?.[0];
    const text = typeof raw === "string" ? raw.trim() : "";

    let current = normalizeBlob(null);
    if (text) {
      try {
        current = normalizeBlob(JSON.parse(text));
      } catch {
        current = normalizeBlob(null);
      }
    }

    const updated = normalizeBlob(updater(current));
    cell.values = [[JSON.stringify(updated)]];
    cell.numberFormat = [["@"]];

    await context.sync();
    return updated;
  });
}

/* =========================
   Sheets
========================= */

async function getVisibleWorksheetNames() {
  if (!globalThis.Excel) throw new Error("Excel.js not available. Are you running inside Excel?");
  await officeReady();

  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name,items/visibility");
    await context.sync();

    const visible = sheets.items
      .filter((ws) => ws.visibility === Excel.SheetVisibility.visible)
      .map((ws) => ws.name);

    visible.sort((a, b) => a.localeCompare(b));
    return visible;
  });
}

async function activateWorksheetByName(sheetName) {
  if (!globalThis.Excel) throw new Error("Excel.js not available. Are you running inside Excel?");
  await officeReady();

  return Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(sheetName);
    ws.activate();
    await context.sync();
  });
}

/* =========================
   User identity
========================= */

function newUuid() {
  if (globalThis.crypto?.randomUUID) return globalThis.crypto.randomUUID();
  const s4 = () => Math.floor((1 + Math.random()) * 0x10000).toString(16).slice(1);
  return `${s4()}${s4()}-${s4()}-${s4()}-${s4()}-${s4()}${s4()}${s4()}`;
}

async function getOrCreateUserKey() {
  if (!globalThis.OfficeRuntime?.storage) return `session-${newUuid()}`;

  let key = await OfficeRuntime.storage.getItem(USER_KEY_STORAGE_KEY);
  if (!key) {
    key = newUuid();
    await OfficeRuntime.storage.setItem(USER_KEY_STORAGE_KEY, key);
  }
  return key;
}

/* =========================
   Settings sheet + user column/blob
========================= */

async function ensureSettingsWorksheet() {
  if (!globalThis.Excel) throw new Error("Excel.js not available. Are you running inside Excel?");
  await officeReady();

  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;

    let settingsSheet;
    try {
      settingsSheet = sheets.getItem(SETTINGS_SHEET_NAME);
      settingsSheet.load("name,visibility");
      await context.sync();
    } catch {
      settingsSheet = sheets.add(SETTINGS_SHEET_NAME);
      settingsSheet.load("name");
      await context.sync();
    }

    settingsSheet.visibility = Excel.SheetVisibility.veryHidden;

    // Signature (optional)
    const a1 = settingsSheet.getRange("A1");
    a1.values = [[SETTINGS_SHEET_NAME]];
    a1.numberFormat = [["@"]];

    await context.sync();
  });
}

async function ensureUserColumnAndBlob(userKey) {
  if (!globalThis.Excel) throw new Error("Excel.js not available. Are you running inside Excel?");
  await officeReady();

  return Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(SETTINGS_SHEET_NAME);

    const used = ws.getUsedRangeOrNullObject();
    used.load("columnCount");
    await context.sync();

    const minCols = 26; // scan at least D..(D+25)
    const colCount = used.isNullObject
      ? USER_COL_START - 1 + minCols
      : Math.max(used.columnCount, USER_COL_START - 1 + minCols);

    const header = ws.getRangeByIndexes(
      USER_KEY_ROW - 1,
      USER_COL_START - 1,
      1,
      colCount - (USER_COL_START - 1)
    );
    header.load("values");
    await context.sync();

    const rowVals = header.values?.[0] ?? [];
    let offset = rowVals.findIndex((v) => String(v ?? "").trim() === userKey);

    if (offset === -1) {
      offset = rowVals.findIndex((v) => String(v ?? "").trim() === "");
      if (offset === -1) offset = rowVals.length;
      header.getCell(0, offset).values = [[userKey]];
    }

    const colIndex0 = USER_COL_START - 1 + offset; // 0-based column index
    const blobCell = ws.getCell(USER_BLOB_ROW - 1, colIndex0);

    blobCell.load("values,address");
    await context.sync();

    const current = blobCell.values?.[0]?.[0];
    const asText = typeof current === "string" ? current.trim() : "";

    let ok = false;
    if (asText) {
      try {
        JSON.parse(asText);
        ok = true;
      } catch {
        ok = false;
      }
    }

    if (!ok) {
      blobCell.values = [[JSON.stringify({ favorites: [], recents: [], options: {} })]];
      blobCell.numberFormat = [["@"]];
    }

    await context.sync();
    return { colIndex0, blobAddress: blobCell.address };
  });
}

/* =========================
   Inventory + Frequency (E)
========================= */

async function syncInventoryPreserveFrequencies(visibleSheetNames, ensureMaxColIndex0) {
  if (!globalThis.Excel) throw new Error("Excel.js not available. Are you running inside Excel?");
  await officeReady();

  const desired = [...visibleSheetNames].sort((a, b) => a.localeCompare(b));

  return Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(SETTINGS_SHEET_NAME);

    const used = ws.getUsedRangeOrNullObject();
    used.load("rowCount,columnCount");
    await context.sync();

    const usedRowCount = used.isNullObject ? 0 : used.rowCount;
    const usedColCount = used.isNullObject ? 0 : used.columnCount;

    const colCount = Math.max(usedColCount, (ensureMaxColIndex0 ?? 0) + 1, USER_COL_START);
    const existingRowCount = Math.max(0, usedRowCount - INVENTORY_START_ROW0);

    let existingNames = [];
    let existingFreq = [];

    if (existingRowCount > 0) {
      const nameRange = ws.getRangeByIndexes(INVENTORY_START_ROW0, INVENTORY_NAME_COL0, existingRowCount, 1);
      nameRange.load("values");

      const freqRange = ws.getRangeByIndexes(
        INVENTORY_START_ROW0,
        USER_COL_START - 1,
        existingRowCount,
        colCount - (USER_COL_START - 1)
      );
      freqRange.load("values");

      await context.sync();

      existingNames = (nameRange.values ?? []).map((r) => String(r?.[0] ?? "").trim());
      existingFreq = freqRange.values ?? [];
    }

    const freqByName = new Map();
    for (let i = 0; i < existingNames.length; i++) {
      const nm = existingNames[i];
      if (!nm) continue;
      freqByName.set(nm, existingFreq[i] ?? []);
    }

    const newRowCount = desired.length;

    if (existingRowCount > 0) {
      const clearRange = ws.getRangeByIndexes(INVENTORY_START_ROW0, 0, existingRowCount, colCount);
      clearRange.clear(Excel.ClearApplyTo.all);
    }

    if (newRowCount > 0) {
      const invRange = ws.getRangeByIndexes(INVENTORY_START_ROW0, 0, newRowCount, INVENTORY_META_COLS);
      invRange.values = desired.map((nm) => [nm, "", ""]);
      // numberFormat must match range dimensions (newRowCount x 3)
      invRange.numberFormat = desired.map(() => ["@", "@", "@"]);
    }

    if (newRowCount > 0) {
      const freqCols = colCount - (USER_COL_START - 1);
      const freqOut = desired.map((nm) => {
        const oldRow = freqByName.get(nm) ?? [];
        const row = new Array(freqCols).fill("");
        for (let c = 0; c < Math.min(freqCols, oldRow.length); c++) {
          const v = oldRow[c];
          row[c] = v === "" || v === null || typeof v === "undefined" ? "" : v;
        }
        return row;
      });

      const freqRangeOut = ws.getRangeByIndexes(
        INVENTORY_START_ROW0,
        USER_COL_START - 1,
        newRowCount,
        colCount - (USER_COL_START - 1)
      );
      freqRangeOut.values = freqOut;
    }

    await context.sync();
    return { sheetCount: newRowCount, colCount };
  });
}

async function ensureInventoryRowAndIncrement(sheetName, userColIndex0, ensureMaxColIndex0) {
  if (!globalThis.Excel) throw new Error("Excel.js not available. Are you running inside Excel?");
  await officeReady();

  return Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(SETTINGS_SHEET_NAME);

    const used = ws.getUsedRangeOrNullObject();
    used.load("rowCount,columnCount");
    await context.sync();

    const usedRowCount = used.isNullObject ? 0 : used.rowCount;
    const usedColCount = used.isNullObject ? 0 : used.columnCount;

    const colCount = Math.max(usedColCount, (ensureMaxColIndex0 ?? 0) + 1, USER_COL_START);
    const existingRowCount = Math.max(0, usedRowCount - INVENTORY_START_ROW0);

    let rowOffset = -1;
    if (existingRowCount > 0) {
      const nameRange = ws.getRangeByIndexes(INVENTORY_START_ROW0, 0, existingRowCount, 1);
      nameRange.load("values");
      await context.sync();

      const names = (nameRange.values ?? []).map((r) => String(r?.[0] ?? "").trim());
      rowOffset = names.findIndex((n) => n === sheetName);
    }

    if (rowOffset === -1) {
      rowOffset = existingRowCount;

      const invRow = ws.getRangeByIndexes(INVENTORY_START_ROW0 + rowOffset, 0, 1, INVENTORY_META_COLS);
      invRow.values = [[sheetName, "", ""]];
      invRow.numberFormat = [["@", "@", "@"]];

      const freqRow = ws.getRangeByIndexes(
        INVENTORY_START_ROW0 + rowOffset,
        USER_COL_START - 1,
        1,
        colCount - (USER_COL_START - 1)
      );
      freqRow.values = [new Array(colCount - (USER_COL_START - 1)).fill("")];

      await context.sync();
    }

    const cell = ws.getCell(INVENTORY_START_ROW0 + rowOffset, userColIndex0);
    cell.load("values");
    await context.sync();

    const current = cell.values?.[0]?.[0];
    const next = toNumberOrZero(current) + 1;
    cell.values = [[next]];

    await context.sync();
    return next;
  });
}

async function deleteInventoryRowIfPresent(sheetName) {
  if (!globalThis.Excel) throw new Error("Excel.js not available. Are you running inside Excel?");
  await officeReady();

  return Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(SETTINGS_SHEET_NAME);

    const used = ws.getUsedRangeOrNullObject();
    used.load("rowCount,columnCount");
    await context.sync();

    if (used.isNullObject) return false;

    const usedRowCount = used.rowCount;
    const usedColCount = used.columnCount;

    const existingRowCount = Math.max(0, usedRowCount - INVENTORY_START_ROW0);
    if (existingRowCount === 0) return false;

    const nameRange = ws.getRangeByIndexes(INVENTORY_START_ROW0, 0, existingRowCount, 1);
    nameRange.load("values");
    await context.sync();

    const names = (nameRange.values ?? []).map((r) => String(r?.[0] ?? "").trim());
    const rowOffset = names.findIndex((n) => n === sheetName);
    if (rowOffset === -1) return false;

    const rowRange = ws.getRangeByIndexes(INVENTORY_START_ROW0 + rowOffset, 0, 1, usedColCount);
    rowRange.delete(Excel.DeleteShiftDirection.up);

    await context.sync();
    return true;
  });
}

/* =========================
   Smart ordering (F)
========================= */

async function readAllFrequenciesForUser(userColIndex0, visibleSheetNames) {
  if (!globalThis.Excel) throw new Error("Excel.js not available. Are you running inside Excel?");
  await officeReady();

  const desiredSet = new Set(visibleSheetNames);

  return Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(SETTINGS_SHEET_NAME);

    const used = ws.getUsedRangeOrNullObject();
    used.load("rowCount");
    await context.sync();

    const usedRowCount = used.isNullObject ? 0 : used.rowCount;
    const existingRowCount = Math.max(0, usedRowCount - INVENTORY_START_ROW0);
    if (existingRowCount === 0) return new Map();

    const nameRange = ws.getRangeByIndexes(INVENTORY_START_ROW0, 0, existingRowCount, 1);
    const freqRange = ws.getRangeByIndexes(INVENTORY_START_ROW0, userColIndex0, existingRowCount, 1);
    nameRange.load("values");
    freqRange.load("values");
    await context.sync();

    const names = (nameRange.values ?? []).map((r) => String(r?.[0] ?? "").trim());
    const freqs = (freqRange.values ?? []).map((r) => r?.[0]);

    const map = new Map();
    for (let i = 0; i < names.length; i++) {
      const nm = names[i];
      if (!nm || !desiredSet.has(nm)) continue;
      map.set(nm, toNumberOrZero(freqs[i]));
    }
    return map;
  });
}

function buildSmartOrder({ visibleNames, favorites, recents, freqMap }) {
  const visibleSet = new Set(visibleNames);

  const fav = (favorites ?? []).filter((n) => visibleSet.has(n));
  const rec = (recents ?? []).filter((n) => visibleSet.has(n));

  const favSet = new Set(fav);
  const recNoFav = rec.filter((n) => !favSet.has(n));
  const recSet = new Set(recNoFav);

  const others = visibleNames.filter((n) => !favSet.has(n) && !recSet.has(n));

  // Significant-frequency ordering: tier desc, then alpha
  others.sort((a, b) => {
    const fa = freqMap?.get(a) ?? 0;
    const fb = freqMap?.get(b) ?? 0;
    const ta = freqTier(fa, 10, 1.35);
    const tb = freqTier(fb, 10, 1.35);
    if (tb !== ta) return tb - ta;
    return a.localeCompare(b);
  });

  return [...fav, ...recNoFav, ...others];
}

/* =========================
   React component
========================= */

const App = (props) => {
  const { title } = props;
  const styles = useStyles();

  // User identity + storage coordinates
  const [userKey, setUserKey] = React.useState(null);
  const [userColIndex0, setUserColIndex0] = React.useState(null);
  const [blobAddress, setBlobAddress] = React.useState(null);
  const [userBlob, setUserBlob] = React.useState(normalizeBlob(null));

  // Frequency map for this user: sheetName -> count
  const [freqMap, setFreqMap] = React.useState(new Map());

  // Sheet list + search
  const [fullList, setFullList] = React.useState([]);
  const [query, setQuery] = React.useState("");
  const [status, setStatus] = React.useState({ state: "idle", message: "" }); // idle|loading|error|info

  // Manage dialog
  const [manageOpen, setManageOpen] = React.useState(false);
  const [manageTab, setManageTab] = React.useState("favorites"); // favorites | settings

  // Manage → Favorites picker search
  const [managePickerQuery, setManagePickerQuery] = React.useState("");

  // Search input ref for focusing
  const searchRef = React.useRef(null);

  const oneDigitActivationEnabled = !!userBlob?.options?.oneDigitActivationEnabled;

  const effectiveQuery = React.useMemo(() => stripLeadingSpaces(query).toLowerCase(), [query]);

  const filteredList = React.useMemo(() => {
    const q = effectiveQuery;
    if (!q) return fullList;
    // Contains match (NOT startsWith)
    return fullList.filter((name) => name.toLowerCase().includes(q));
  }, [fullList, effectiveQuery]);

  const favoritesDisplay = React.useMemo(() => {
    const visibleSet = new Set(fullList);
    return (userBlob.favorites ?? []).filter((n) => visibleSet.has(n));
  }, [userBlob.favorites, fullList]);

  const smartOrdered = React.useMemo(() => {
    return buildSmartOrder({
      visibleNames: filteredList,
      favorites: userBlob.favorites,
      recents: userBlob.recents,
      freqMap,
    });
  }, [filteredList, userBlob.favorites, userBlob.recents, freqMap]);

  const isReady = status.state === "idle" && userColIndex0 !== null && !!blobAddress;

  /* ===== Favorites mutations (append to end) ===== */

  const addFavoriteAppend = React.useCallback(
    async (sheetName) => {
      if (!blobAddress) return;
      const updated = await updateUserBlobAtAddress(blobAddress, (b) => {
        const fav = b.favorites ?? [];
        if (fav.includes(sheetName)) return b;
        const next = [...fav, sheetName].slice(0, MAX_FAVORITES);
        return { ...b, favorites: next };
      });
      setUserBlob(updated);
    },
    [blobAddress]
  );

  const removeFavorite = React.useCallback(
    async (sheetName) => {
      if (!blobAddress) return;
      const updated = await updateUserBlobAtAddress(blobAddress, (b) => ({
        ...b,
        favorites: (b.favorites ?? []).filter((x) => x !== sheetName),
      }));
      setUserBlob(updated);
    },
    [blobAddress]
  );

  const setFavoritesOrder = React.useCallback(
    async (nextFavorites) => {
      if (!blobAddress) return;
      const updated = await updateUserBlobAtAddress(blobAddress, (b) => ({
        ...b,
        favorites: (nextFavorites ?? []).slice(0, MAX_FAVORITES),
      }));
      setUserBlob(updated);
    },
    [blobAddress]
  );

  /* ===== Centralized cleanup for missing sheets ===== */

  const handleMissingSheet = React.useCallback(
    async (sheetName) => {
      try {
        // 1) Remove from favorites + recents blob
        if (blobAddress) {
          const updated = await updateUserBlobAtAddress(blobAddress, (b) => ({
            ...b,
            favorites: (b.favorites ?? []).filter((x) => x !== sheetName),
            recents: (b.recents ?? []).filter((x) => x !== sheetName),
          }));
          setUserBlob(updated);
        }

        // 2) Remove inventory row (row 52+)
        try {
          await deleteInventoryRowIfPresent(sheetName);
        } catch {
          // ignore; still proceed with UI cleanup
        }

        // 3) Remove from UI lists + freqMap
        setFullList((prev) => prev.filter((x) => x !== sheetName));
        setFreqMap((prev) => {
          const m = new Map(prev);
          m.delete(sheetName);
          return m;
        });

        // 4) Notify user
        setStatus({
          state: "info",
          message: `Sheet "${sheetName}" no longer exists. Removed from favorites/recents and cleaned up.`,
        });
      } catch (err) {
        setStatus({ state: "error", message: err?.message || String(err) });
      }
    },
    [blobAddress]
  );

  /* ===== Activation pipeline ===== */

  const activateOrCleanup = React.useCallback(
  async (sheetName) => {
    if (!isReady) return;

    // Pre-check: if it's not in our current visible list, treat as missing immediately.
    // This prevents Excel.run from throwing and triggering dev overlays.
    if (!fullList.includes(sheetName)) {
      await handleMissingSheet(sheetName);
      return;
    }

    try {
      await activateWorksheetByName(sheetName);

      const updated = await updateUserBlobAtAddress(blobAddress, (b) => ({
        ...b,
        recents: pushUniqueFront(b.recents ?? [], sheetName, MAX_RECENTS),
      }));
      setUserBlob(updated);

      const nextVal = await ensureInventoryRowAndIncrement(sheetName, userColIndex0, userColIndex0);
      setFreqMap((prev) => {
        const m = new Map(prev);
        m.set(sheetName, nextVal);
        return m;
      });
      try {
        await hideTaskpaneSafely();
      } catch {
        // TEMP/Best-effort: ignore hide failures; sheet activation already succeeded.
      }
    } catch (err) {
      if (isMissingSheetError(err)) {
        await handleMissingSheet(sheetName);
        return;
      }
      const msg = err?.message || String(err);
      setStatus({ state: "error", message: msg });
    }
  },
  [isReady, fullList, blobAddress, userColIndex0, handleMissingSheet]
);


  /* ===== Init ===== */

  React.useEffect(() => {
    let cancelled = false;

    (async () => {
      try {
        setStatus({ state: "loading", message: "Initializing…" });

        // TEMP: enable DEBUG_SHARED_RUNTIME_PROBE above to run Shared Runtime probes.
        installSharedRuntimeProbeOnce();

        await ensureSettingsWorksheet();

        const key = await getOrCreateUserKey();
        const { colIndex0, blobAddress: addr } = await ensureUserColumnAndBlob(key);

        if (cancelled) return;
        setUserKey(key);
        setUserColIndex0(colIndex0);
        setBlobAddress(addr);

        const initialBlob = await readUserBlobAtAddress(addr);
        if (cancelled) return;
        setUserBlob(initialBlob);

        const names = await getVisibleWorksheetNames();
        if (cancelled) return;

        await syncInventoryPreserveFrequencies(names, colIndex0);

        const fm = await readAllFrequenciesForUser(colIndex0, names);
        if (cancelled) return;
        setFreqMap(fm);

        setFullList(names);
        setStatus({ state: "idle", message: "" });

        // Focus search box on open
        setTimeout(() => {
          try {
            searchRef.current?.focus?.();
          } catch {
            /* ignore */
          }
        }, 0);
      } catch (err) {
        if (cancelled) return;
        setStatus({ state: "error", message: err?.message || String(err) });
      }
    })();

    return () => {
      cancelled = true;
    };
  }, []);

  /* ===== Search box: one-digit favorites activation ===== */

  const onSearchKeyDown = React.useCallback(
    async (e) => {
      if (!oneDigitActivationEnabled) return;
      if (!isReady) return;

      // If user already typed something, do nothing (normal search)
      if (query.length !== 0) return;

      // If modifiers are pressed, treat as normal typing
      if (e.ctrlKey || e.altKey || e.metaKey) return;

      const key = e.key;

      // Leading-space "escape": allow space to be typed; filtering ignores it via stripLeadingSpaces()
      if (key === " ") return;

      // Only consider digits 0-9
      if (!/^[0-9]$/.test(key)) return;

      const idx1 = digitToFavoriteIndex1Based(key);
      if (!idx1) return;

      const favName = favoritesDisplay[idx1 - 1];

      // If not enough favorites, treat as normal search (digit should appear)
      if (!favName) return;

      e.preventDefault();
      e.stopPropagation();

      await activateOrCleanup(favName);
    },
    [oneDigitActivationEnabled, isReady, query, favoritesDisplay, activateOrCleanup]
  );

  /* ===== Overflow menu helpers ===== */

  const openManage = React.useCallback(() => {
    setManageTab("favorites");
    setManageOpen(true);
    // Don’t reset picker query automatically; preserve as user works
  }, []);

  /* ===== Manage dialog: DnD favorites ===== */

  const [dragFrom, setDragFrom] = React.useState(null);

  const onFavDragStart = React.useCallback((index) => {
    setDragFrom(index);
  }, []);

  const onFavDragOver = React.useCallback((e) => {
    // Needed so onDrop fires
    e.preventDefault();
  }, []);

  const onFavDrop = React.useCallback(
    async (dropIndex) => {
      if (dragFrom === null || dragFrom === dropIndex) return;
      const next = moveArrayItem(favoritesDisplay, dragFrom, dropIndex);
      setDragFrom(null);
      await setFavoritesOrder(next);
    },
    [dragFrom, favoritesDisplay, setFavoritesOrder]
  );

  /* ===== Settings toggle ===== */

  const setOneDigitActivationEnabled = React.useCallback(
    async (enabled) => {
      if (!blobAddress) return;
      const updated = await updateUserBlobAtAddress(blobAddress, (b) => ({
        ...b,
        options: { ...(b.options ?? {}), oneDigitActivationEnabled: !!enabled },
      }));
      setUserBlob(updated);
    },
    [blobAddress]
  );

  const isFavorite = React.useCallback((name) => (userBlob.favorites ?? []).includes(name), [userBlob.favorites]);

  const RowMenu = ({ sheetName }) => {
    const fav = isFavorite(sheetName);
    return (
      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <Button
            appearance="subtle"
            icon={<MoreHorizontal24Regular />}
            onClick={(e) => {
              // Don’t activate row
              e.stopPropagation();
            }}
            aria-label="More actions"
          />
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            {!fav ? (
              <MenuItem
                onClick={async () => {
                  await addFavoriteAppend(sheetName);
                }}
              >
                Add to favorites
              </MenuItem>
            ) : (
              <MenuItem
                onClick={async () => {
                  await removeFavorite(sheetName);
                }}
              >
                Remove from favorites
              </MenuItem>
            )}
            <MenuItem
              onClick={() => {
                openManage();
              }}
            >
              Manage…
            </MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>
    );
  };

  RowMenu.propTypes = { sheetName: PropTypes.string.isRequired };

  /* ===== Manage → Favorites: All sheets picker ===== */

  const managePickerEffectiveQuery = React.useMemo(() => stripLeadingSpaces(managePickerQuery).toLowerCase(), [
    managePickerQuery,
  ]);

  const allSheetsForPicker = React.useMemo(() => {
    const base = [...fullList].sort((a, b) => a.localeCompare(b));
    const q = managePickerEffectiveQuery;
    if (!q) return base;
    return base.filter((nm) => nm.toLowerCase().includes(q));
  }, [fullList, managePickerEffectiveQuery]);

  const canAddMoreFavorites = favoritesDisplay.length < MAX_FAVORITES;

  const toggleFavoriteFromPicker = React.useCallback(
    async (sheetName) => {
      const fav = isFavorite(sheetName);

      if (fav) {
        await removeFavorite(sheetName);
        return;
      }

      if (!canAddMoreFavorites) {
        setStatus({
          state: "info",
          message: `Favorites are limited to ${MAX_FAVORITES}. Remove one before adding another.`,
        });
        return;
      }

      await addFavoriteAppend(sheetName);
    },
    [isFavorite, removeFavorite, addFavoriteAppend, canAddMoreFavorites]
  );

  return (
    <div className={styles.root}>
      <div style={{ padding: "12px 16px" }}>
        <Text weight="semibold" size={600}>
          JumpTo
        </Text>
      </div>

      <div className={styles.section}>
        <div className={styles.manageHeader}>
          <Button
            appearance="subtle"
            onClick={() => {
              openManage();
            }}
          >
            Manage…
          </Button>
        </div>

        <div style={{ marginTop: 8 }}>
          <Input
            ref={searchRef}
            value={query}
            onChange={(_, data) => setQuery(data.value)}
            onKeyDown={onSearchKeyDown}
            placeholder="Type to filter sheets…"
            appearance="outline"
            disabled={status.state === "loading"}
          />
        </div>

        {status.state === "loading" && <div className={styles.meta}>{status.message}</div>}
        {status.state === "error" && (
          <div className={styles.meta} style={{ color: "crimson" }}>
            Error: {status.message}
          </div>
        )}
        {status.state === "info" && <div className={styles.meta}>{status.message}</div>}

        {/* Favorites section */}
        {favoritesDisplay.length > 0 ? (
          <>
            <Text className={styles.subheader} weight="semibold">
              Favorites
            </Text>
            <div className={styles.list}>
              {favoritesDisplay.map((name, i) => (
                <div
                  key={`fav:${name}`}
                  className={styles.row}
                  onClick={() => activateOrCleanup(name)}
                  title={`Activate favorite "${name}"`}
                  aria-disabled={!isReady}
                  style={!isReady ? { opacity: 0.6, pointerEvents: "none" } : undefined}
                >
                  <div className={styles.rowLeft}>
                    <span className={styles.hotkey}>{FAVORITE_HOTKEY_LABELS[i] ?? ""}</span>
                    <span className={styles.sheetName}>{name}</span>
                  </div>
                  <div className={styles.rowRight}>
                    <RowMenu sheetName={name} />
                  </div>
                </div>
              ))}
              {oneDigitActivationEnabled ? (
                <div className={styles.meta}>
                  Tip: with an empty search box, type 1–9 (0=10) to open a favorite. To search starting with digits, begin with a space.
                </div>
              ) : null}
            </div>
          </>
        ) : null}

        {/* Main list (smart ordered) */}
        <div className={styles.list}>
          {smartOrdered.map((name) => (
            <div
              key={name}
              className={styles.row}
              onClick={() => activateOrCleanup(name)}
              title={`Activate "${name}"`}
              aria-disabled={!isReady}
              style={!isReady ? { opacity: 0.6, pointerEvents: "none" } : undefined}
            >
              <div className={styles.rowLeft}>
                <span style={{ width: "18px" }} />
                <span className={styles.sheetName}>{name}</span>
              </div>
              <div className={styles.rowRight}>
                <RowMenu sheetName={name} />
              </div>
            </div>
          ))}
        </div>

        {/* =========================
            Manage Dialog
        ========================= */}
        <Dialog open={manageOpen} onOpenChange={(_, data) => setManageOpen(data.open)}>
          <DialogSurface>
            <DialogBody>
              <DialogTitle
                action={
                  <Button
                    appearance="subtle"
                    icon={<Dismiss24Regular />}
                    onClick={() => setManageOpen(false)}
                    aria-label="Close"
                  />
                }
              >
                Manage
              </DialogTitle>

              <DialogContent>
                <TabList selectedValue={manageTab} onTabSelect={(_, data) => setManageTab(String(data.value))} appearance="subtle">
                  <Tab value="favorites">Favorites</Tab>
                  <Tab value="settings">Settings</Tab>
                </TabList>

                {manageTab === "favorites" ? (
                  <div style={{ marginTop: 10 }}>
                    <Text weight="semibold">Reorder favorites</Text>
                    <div className={styles.meta}>Drag and drop to reorder. (Adding favorites appends to the end.)</div>

                    {favoritesDisplay.length === 0 ? (
                      <div className={styles.meta}>No favorites yet. Use the picker below (or the “…” menu) to add some.</div>
                    ) : (
                      <div>
                        {favoritesDisplay.map((name, idx) => (
                          <div
                            key={`dnd:${name}`}
                            className={styles.favoritesDnDRow}
                            draggable
                            onDragStart={() => onFavDragStart(idx)}
                            onDragOver={onFavDragOver}
                            onDrop={() => onFavDrop(idx)}
                            title="Drag to reorder"
                          >
                            <span className={styles.dragHandle} aria-hidden="true">
                              <ReOrderDotsVertical24Regular />
                            </span>
                            <span className={styles.dndName}>
                              {FAVORITE_HOTKEY_LABELS[idx] ?? ""} {FAVORITE_HOTKEY_LABELS[idx] ? "—" : ""} {name}
                            </span>
                            <Button
                              appearance="subtle"
                              onClick={async (e) => {
                                e.preventDefault();
                                e.stopPropagation();
                                await removeFavorite(name);
                              }}
                            >
                              Remove
                            </Button>
                          </div>
                        ))}
                      </div>
                    )}

                    {/* All sheets picker (NEW) */}
                    <div className={styles.pickerBox}>
                      <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", gap: 12 }}>
                        <Text weight="semibold">All sheets</Text>
                        <Text className={styles.meta}>
                          {favoritesDisplay.length}/{MAX_FAVORITES} favorites
                        </Text>
                      </div>

                      <div style={{ marginTop: 8 }}>
                        <Input
                          value={managePickerQuery}
                          onChange={(_, data) => setManagePickerQuery(data.value)}
                          placeholder="Search sheets…"
                          appearance="outline"
                          // IMPORTANT: do not focus-steal from main search unless user clicks here
                        />
                      </div>

                      {!canAddMoreFavorites ? (
                        <div className={styles.meta}>
                          Favorites limit reached. Uncheck a favorite below (or remove above) to add another.
                        </div>
                      ) : null}

                      <div className={styles.pickerList} role="list" aria-label="All sheets picker">
                        {allSheetsForPicker.length === 0 ? (
                          <div className={styles.meta} style={{ padding: "10px 2px" }}>
                            No matching sheets.
                          </div>
                        ) : (
                          allSheetsForPicker.map((nm) => {
                            const fav = isFavorite(nm);
                            const disabled = !fav && !canAddMoreFavorites;

                            return (
                              <div key={`pick:${nm}`} className={styles.pickerRow} role="listitem">
                                <span className={styles.pickerName} title={nm}>
                                  {nm}
                                </span>

                                <Checkbox
                                  checked={fav}
                                  disabled={disabled}
                                  label={fav ? "Favorite" : ""}
                                  onChange={async (e) => {
                                    // No activation in manage dialog; only toggling favorites
                                    e.preventDefault();
                                    e.stopPropagation();
                                    await toggleFavoriteFromPicker(nm);
                                  }}
                                />
                              </div>
                            );
                          })
                        )}
                      </div>

                      <div className={styles.meta}>
                        Tip: toggling here only adds/removes favorites. To open a sheet, close Manage and click it in the main list (or use
                        digit activation).
                      </div>
                    </div>
                  </div>
                ) : null}

                {manageTab === "settings" ? (
                  <div style={{ marginTop: 10 }}>
                    <Text weight="semibold">Per-user workbook settings</Text>

                    <div style={{ marginTop: 12, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                      <div>
                        <Text>One-digit favorite activation</Text>
                        <div className={styles.meta}>
                          When enabled: with an empty search box, typing 1–9 (0=10) activates that favorite. To search starting with
                          digits, start with a space.
                        </div>
                      </div>

                      <Switch checked={oneDigitActivationEnabled} onChange={(_, data) => setOneDigitActivationEnabled(!!data.checked)} />
                    </div>

                    <div style={{ marginTop: 16 }}>
                      <Text weight="semibold">Coming next</Text>
                      <div className={styles.meta}>
                        FullList rows to display • Favorites rows to display • Recents max (global) • Other future settings
                      </div>
                    </div>

                    {/* Iteration 36: closeContainer viability check (dev aid). */}
                    <div style={{ marginTop: 20 }}>
                      <Text weight="semibold">Diagnostics</Text>
                      <div className={styles.meta}>
                        Test whether this Excel build can close the taskpane programmatically. This should be treated as a last step
                        after work is done.
                      </div>
                      <div style={{ marginTop: 8 }}>
                        <Button
                          onClick={async () => {
                            try {
                              await hideTaskpaneSafely();
                            } catch {
                              // ignore
                            }
                          }}
                        >
                          Close taskpane now (test)
                        </Button>
                      </div>
                    </div>
                  </div>
                ) : null}
              </DialogContent>

              <DialogActions>
                <Button
                  appearance="primary"
                  onClick={() => {
                    setManageOpen(false);
                  }}
                >
                  Done
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      </div>
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
