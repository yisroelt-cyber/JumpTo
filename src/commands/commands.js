// DEBUG: persistence instrumentation (temporary)
const DEBUG_PERSIST = true;



function dbgPersist(tag, payload) {
  if (!DEBUG_PERSIST) return;
  try {
    console.groupCollapsed(`[JumpTo Persist DEBUG] ${tag}`);


function sendPersistDiagToDialog(tag, payload) {
  if (!DEBUG_PERSIST) return;
  try {
    if (typeof Office !== "undefined" && Office.context && Office.context.ui && Office.context.ui.messageParent) {
      Office.context.ui.messageParent(JSON.stringify({ type: "persistDiag", tag, payload }));
    }
  } catch (e) {
    // no-op
  }
}
    console.log(payload);
    console.groupEnd();
  } catch (e) {
    // no-op
  }
}
/*
  commands.js – Option B engine with cache + refresh-on-open (signature based)
*/

import {
  getJumpToState,
  toggleFavorite as toggleFavoriteInStorage,
  setFavorites as setFavoritesInStorage,
  recordActivation,
  setUiSettings as setUiSettingsInStorage,
} from "../services/jumpToStorage";

let lockBusy = false;
const lockQueue = [];
const pendingStateRequests = [];

let cachedState = null;
let cachedSignature = "";
let lastCheckTs = 0;
const CHECK_TTL_MS = 1500;

// Global (per-user) option keys stored in OfficeRuntime.storage
const OPT_ROW_HEIGHT = "JumpTo.Option.RowHeightPreset";
const OPT_BASELINE_ORDER = "JumpTo.Option.BaselineOrder";
const OPT_FREQUENT_ON_TOP = "JumpTo.Option.FrequentOnTop";
const OPT_FAV_PERCENT = "JumpTo.Option.FavPercentManual";
const OPT_RECENTS_DISPLAY_COUNT = "JumpTo.Option.RecentsDisplayCount";

// Legacy key (previously global) for one-digit activation; now workbook-scoped.
const OPT_ONE_DIGIT_LEGACY = "JumpTo.Option.OneDigitActivation";


function withLock(fn) {
  return new Promise((resolve, reject) => {
    lockQueue.push({ fn, resolve, reject });
    pump();
  });
}

async function pump() {
  if (lockBusy || lockQueue.length === 0) return;
  lockBusy = true;
  const job = lockQueue.shift();
  try {
    const result = await job.fn();
    job.resolve(result);
  } catch (e) {
    job.reject(e);
  } finally {
    lockBusy = false;
    pump();
  }
}

async function computeSheetSignature() {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/id,name,visibility");
    await context.sync();
    return sheets.items
      .filter((s) => s.visibility === "visible")
      .map((s) => `${s.id}:${s.name}`)
      .join("|");
  });
}

async function ensureFreshState() {
  const now = Date.now();
  if (now - lastCheckTs < CHECK_TTL_MS) return false;
  lastCheckTs = now;

  const sig = await computeSheetSignature();
  if (sig === cachedSignature && cachedState) return false;

  cachedState = await getJumpToState();
  cachedSignature = sig;
  return true;
}


async function getActiveWorksheetId() {
  return Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getActiveWorksheet();
    ws.load("id");
    await context.sync();
    return ws.id;
  });
}

async function buildDialogState(baseState) {

// DEBUG: persist diagnostics (temporary)
try {
  sendPersistDiag("env", {
    href: (typeof window !== "undefined" && window.location) ? window.location.href : null,
    origin: (typeof window !== "undefined" && window.location) ? window.location.origin : null,
    hasOfficeRuntime: typeof OfficeRuntime !== "undefined",
    hasOfficeRuntimeStorage: (typeof OfficeRuntime !== "undefined") && !!(OfficeRuntime.storage && OfficeRuntime.storage.getItem),
  });
} catch (e) {}

async function dbgGetPersistKey(key) {
  try {
    const v = await OfficeRuntime.storage.getItem(key);
    sendPersistDiag(`getItem ${key}`, { value: v });
    return v;
  } catch (e) {
    sendPersistDiag(`getItem ERROR ${key}`, { error: String(e) });
    throw e;
  }
}



// --- Persist debug: environment + tracing (Excel restart diagnosis) ---
const PERSIST_TRACE_MARK = true;
const PERSIST_READ_KEYS = {
  rowHeight: "JumpTo.Option.RowHeightPreset",
  favPercent: "JumpTo.Option.FavPercentManual",
  recentsCount: "JumpTo.Option.RecentsDisplayCount",
  baselineOrder: "JumpTo.Option.BaselineOrder",
  frequentOnTop: "JumpTo.Option.FrequentOnTop",
};

const env = {
  href: (typeof window !== "undefined" && window.location) ? window.location.href : null,
  origin: (typeof window !== "undefined" && window.location) ? window.location.origin : null,
  hasOfficeRuntime: typeof OfficeRuntime !== "undefined",
  hasOfficeRuntimeStorage: (typeof OfficeRuntime !== "undefined") && !!(OfficeRuntime.storage && OfficeRuntime.storage.getItem),
};
dbgPersist("env", env);
sendPersistDiagToDialog("env", env);

function trace(tag, payload) {
  dbgPersist(tag, payload);
  sendPersistDiagToDialog(tag, payload);
}

async function dbgGet(key) {
  try {
    const v = await OfficeRuntime.storage.getItem(key);
    trace(`getItem ${key}`, { value: v });
    return v;
  } catch (e) {
    trace(`getItem ERROR ${key}`, { error: String(e) });
    throw e;
  }
}

  if (!baseState) trace("earlyReturn baseState", { baseState });
      return baseState;

  const activeId = await getActiveWorksheetId();

  const sheetsArr = Array.isArray(baseState.sheets) ? baseState.sheets : [];
  const visibleIds = new Set(sheetsArr.map((s) => s.id));
  const idToName = new Map(sheetsArr.map((s) => [s.id, s.name]));

  const baseRecents = Array.isArray(baseState.recents) ? baseState.recents : [];
  const recentIds = baseRecents
    .map((r) => (typeof r === "string" ? r : r?.id))
    .filter(Boolean);

  // Defaults (legacy fallbacks)
  let rowHeightPreset = "Standard";

  // Workbook-scoped: one-digit activation (default ON if unset)
  let oneDigitActivationEnabled = baseState.settings?.oneDigitActivationEnabled;
  if (oneDigitActivationEnabled === undefined) oneDigitActivationEnabled = true;

  // Global-scoped settings (legacy fallback from workbook settings blob)
  let baselineOrder = String(baseState.settings?.baselineOrder || "workbook");
  let frequentOnTop =
    baseState.settings?.frequentOnTop === undefined
      ? true
      : !!baseState.settings?.frequentOnTop;
  let favPercentManual = Number.isFinite(Number(baseState.settings?.favPercentManual))
    ? Number(baseState.settings?.favPercentManual)
    : 50;
  let recentsDisplayCount = Number.isFinite(Number(baseState.settings?.recentsDisplayCount))
    ? Number(baseState.settings?.recentsDisplayCount)
    : 20;

  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage?.getItem) {
      const v = await OfficeRuntime.storage.getItem(OPT_ROW_HEIGHT);
      if (v) rowHeightPreset = String(v);

      const bo = await OfficeRuntime.storage.getItem(OPT_BASELINE_ORDER);
      if (bo) baselineOrder = String(bo);

      const fot = await OfficeRuntime.storage.getItem(OPT_FREQUENT_ON_TOP);
      if (fot === "false") frequentOnTop = false;
      else if (fot === "true") frequentOnTop = true;

      const fp = await OfficeRuntime.storage.getItem(OPT_FAV_PERCENT);
      if (fp !== null && fp !== undefined && fp !== "") favPercentManual = Number(fp);

      const rc = await OfficeRuntime.storage.getItem(OPT_RECENTS_DISPLAY_COUNT);
      if (rc !== null && rc !== undefined && rc !== "") recentsDisplayCount = Number(rc);

      // Legacy one-digit activation was global; if workbook doesn't yet have an override, seed from legacy value.
      if (baseState.settings?.oneDigitActivationEnabled === undefined) {
        const od = await OfficeRuntime.storage.getItem(OPT_ONE_DIGIT_LEGACY);
        if (od === "false") oneDigitActivationEnabled = false;
        else if (od === "true") oneDigitActivationEnabled = true;
      }

      // Best-effort migration: promote legacy workbook-stored globals into global storage if missing.
      if (!bo && baseState.settings?.baselineOrder)
        await OfficeRuntime.storage.setItem(
          OPT_BASELINE_ORDER,
          String(baseState.settings.baselineOrder)
        );

      if (fot === null || fot === undefined) {
        if (baseState.settings?.frequentOnTop !== undefined)
          await OfficeRuntime.storage.setItem(
            OPT_FREQUENT_ON_TOP,
            baseState.settings.frequentOnTop ? "true" : "false"
          );
      }

      if (
        (fp === null || fp === undefined || fp === "") &&
        baseState.settings?.favPercentManual !== undefined
      )
        await OfficeRuntime.storage.setItem(
          OPT_FAV_PERCENT,
          String(baseState.settings.favPercentManual)
        );

      if (
        (rc === null || rc === undefined || rc === "") &&
        baseState.settings?.recentsDisplayCount !== undefined
      )
        await OfficeRuntime.storage.setItem(
          OPT_RECENTS_DISPLAY_COUNT,
          String(baseState.settings.recentsDisplayCount)
        );
    }
  } catch {
    // ignore
  }

  // Clamp recentsDisplayCount for use in filtered Recents list.
  const n = Number.isFinite(recentsDisplayCount)
    ? Math.max(1, Math.min(20, Math.floor(recentsDisplayCount)))
    : 20;

  const filtered = [];
  for (const id of recentIds) {
    if (id === activeId) continue;
    if (!visibleIds.has(id)) continue;
    filtered.push(id);
    if (filtered.length >= n) break;
  }

  return {
    ...baseState,
    // Keep workbook settings minimal; dialog UI can still display global values (provided via `global`).
    settings: { favPercentManual, recentsDisplayCount },
    global: { oneDigitActivationEnabled, rowHeightPreset, baselineOrder, frequentOnTop },
    recents: filtered.map((id) => ({ id, name: idToName.get(id) || "" })),
  };
}



async function setGlobalUiSettings(patch) {
  const p = patch && typeof patch === "object" ? patch : {};
  if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage?.setItem) return;

  // Only persist recognized global-scoped keys.
  const writes = [];
  if (p.favPercentManual !== undefined) writes.push(OfficeRuntime.storage.setItem(OPT_FAV_PERCENT, String(p.favPercentManual)));
  if (p.recentsDisplayCount !== undefined) writes.push(OfficeRuntime.storage.setItem(OPT_RECENTS_DISPLAY_COUNT, String(p.recentsDisplayCount)));
  if (p.baselineOrder !== undefined) writes.push(OfficeRuntime.storage.setItem(OPT_BASELINE_ORDER, String(p.baselineOrder)));
  if (p.frequentOnTop !== undefined) writes.push(OfficeRuntime.storage.setItem(OPT_FREQUENT_ON_TOP, p.frequentOnTop ? "true" : "false"));
  await Promise.all(writes);
}

async function activateSheetById(sheetId) {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/id,name");
    await context.sync();
    const ws = sheets.items.find((s) => s.id === sheetId);
    if (!ws) throw new Error("Sheet not found");
    context.workbook.worksheets.getItem(ws.name).activate();
    await context.sync();
  });
}

function openJumpDialog(event) {
  const dialogUrl = new URL("./dialog.html", window.location.href).toString();

  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 70, width: 45, displayInIframe: true },
    (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        event.completed();
        return;
      }

      const dialog = result.value;

      const reply = (msg) => {
        try {
          dialog.messageChild(JSON.stringify(msg));
        } catch {}
      };


// DEBUG: persistence diagnostics sender (temporary)
const sendPersistDiag = (tag, payload) => {
  try { reply({ type: "persistDiag", tag, payload }); } catch {}
};

async function dbgSetPersistKey(key, value) {
  sendPersistDiag(`setItem ${key}`, { value });
  return OfficeRuntime.storage.setItem(key, value);
}


      const flushStateQueue = async () => {
        // Phase 4: fast-first render. If we have an in-memory cache, use it immediately.
        // Otherwise, try a perf-cache-backed state (preferCache) before doing the full refresh.
        if (cachedState) {
          const state = await buildDialogState(cachedState);
          while (pendingStateRequests.length) {
            pendingStateRequests.pop();
            reply({ type: "stateData", state });
          }
        } else {
          try {
            cachedState = await getJumpToState({ preferCache: true });
            if (cachedState) {
              const state = await buildDialogState(cachedState);
              while (pendingStateRequests.length) {
                pendingStateRequests.pop();
                reply({ type: "stateData", state });
              }
            }
          } catch {
            // ignore; fall through to full refresh
          }
        }

        const changed = await ensureFreshState();
        if (changed && cachedState) {
          const state = await buildDialogState(cachedState);
          reply({ type: "stateData", state });
        }
      };

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
        let msg;
        try {
          msg = JSON.parse(arg.message);
        } catch {
          return;
        }

        if (msg.type === "ping") {
          reply({ type: "parentReady" });
          return;
        }

        if (msg.type === "getSheets") {
          pendingStateRequests.push(true);
          await withLock(flushStateQueue);
          return;
        }

        if (msg.type === "toggleFavorite") {
          await withLock(async () => {
            await toggleFavoriteInStorage(msg.sheetId);
            cachedState = await getJumpToState();
            const state = await buildDialogState(cachedState);
            reply({ type: "stateData", state });
          });
          return;
        }

        if (msg.type === "setFavorites") {
          const ids = Array.isArray(msg.favorites) ? msg.favorites.filter(Boolean) : [];
          await withLock(async () => {
            await setFavoritesInStorage(ids);
            if (!cachedState) {
              cachedState = await getJumpToState();
            } else {
              const idToName = new Map((cachedState.sheets || []).map((s) => [s.id, s.name]));
              cachedState = {
                ...cachedState,
                favorites: ids.slice(0, 20).map((id) => ({ id, name: idToName.get(id) || "" })),
              };
            }
            const state = await buildDialogState(cachedState);
            reply({ type: "stateData", state });
          });
          return;
        }

        if (msg.type === "setUiSettings") {
          const patch = msg.settings && typeof msg.settings === "object" ? msg.settings : {};
          await withLock(async () => {
            await setGlobalUiSettings(patch);
            cachedState = await getJumpToState();
            const state = await buildDialogState(cachedState);
            reply({ type: "stateData", state });
          });
          return;
        }


if (msg.type === "setRowHeightPreset") {
  const preset = typeof msg.preset === "string" ? msg.preset : "";
  if (!preset) return;
  await withLock(async () => {
    try {
      if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage?.setItem) {
        await dbgSetPersistKey("JumpTo.Option.RowHeightPreset", preset);
      }
    } catch {}
    cachedState = await getJumpToState();
    const state = await buildDialogState(cachedState);
    reply({ type: "stateData", state });
  });
  return;
}

        if (msg.type === "setOneDigitActivation") {
          const enabled = !!msg.enabled;
          await withLock(async () => {
            // Workbook-scoped: persist in workbook settings blob.
            await setUiSettingsInStorage({ oneDigitActivationEnabled: enabled });
            if (!cachedState) {
              cachedState = await getJumpToState();
            } else {
              cachedState = {
                ...cachedState,
                settings: { ...(cachedState.settings || {}), oneDigitActivationEnabled: enabled },
              };
            }
            const state = await buildDialogState(cachedState);
            reply({ type: "stateData", state });
          });
          return;
        }

if (msg.type === "selectSheet") {
          const sheetId = msg.sheetId;

          // Snapshot-based persistence: the dialog may close immediately after selection,
          // so carry the latest state in the select message and persist it from the parent
          // *after* the sheet activation has been initiated.
          const snapshot = msg.snapshot && typeof msg.snapshot === "object" ? msg.snapshot : {};
          const uiSettings = snapshot.uiSettings && typeof snapshot.uiSettings === "object" ? snapshot.uiSettings : null;
          const favorites = Array.isArray(snapshot.favorites) ? snapshot.favorites.filter(Boolean) : null;
          const rowHeightPreset = typeof snapshot.rowHeightPreset === "string" ? snapshot.rowHeightPreset : "";

          // Close + complete immediately so the dialog feels instant.
          try {
            dialog.close();
          } catch {}
          event.completed();

          // Continue work in the background so UI close is not blocked by Excel writes.
          (async () => {
            await withLock(async () => {
              if (sheetId) {
                await activateSheetById(sheetId);
                await recordActivation(sheetId);
              }

              // Persist latest state AFTER activation so persistence work doesn't delay the jump.
              if (rowHeightPreset) {
                try {
                  if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage?.setItem) {
                    await dbgSetPersistKey("JumpTo.Option.RowHeightPreset", rowHeightPreset);
                  }
                } catch {}
              }

              const oneDigitActivationEnabled = !!snapshot.oneDigitActivationEnabled;

              try {
                if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage?.setItem) {
                  await OfficeRuntime.storage.setItem(
                    "JumpTo.Option.OneDigitActivation",
                    oneDigitActivationEnabled ? "true" : "false"
                  );
                }
              } catch {}

              if (uiSettings) {
                await setUiSettingsInStorage(uiSettings);
              }

              if (favorites) {
                await setFavoritesInStorage(favorites);
              }

              // Keep cache coherent for the next dialog open.
              cachedState = await getJumpToState();
            });
          })().catch((err) => console.error("selectSheet background handler failed:", err));

          return;
        }

        if (msg.type === "cancel") {
          const snapshot = msg.snapshot && typeof msg.snapshot === "object" ? msg.snapshot : {};
          const uiSettings = snapshot.uiSettings && typeof snapshot.uiSettings === "object" ? snapshot.uiSettings : null;
          const favorites = Array.isArray(snapshot.favorites) ? snapshot.favorites.filter(Boolean) : null;
          const rowHeightPreset = typeof snapshot.rowHeightPreset === "string" ? snapshot.rowHeightPreset : "";

          try {
            dialog.close();
          } catch {}
          event.completed();

          (async () => {
            await withLock(async () => {
              if (rowHeightPreset) {
                try {
                  if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage?.setItem) {
                    await dbgSetPersistKey("JumpTo.Option.RowHeightPreset", rowHeightPreset);
                  }
                } catch {}
              }

              const oneDigitActivationEnabled = !!snapshot.oneDigitActivationEnabled;

              try {
                if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage?.setItem) {
                  await OfficeRuntime.storage.setItem(
                    "JumpTo.Option.OneDigitActivation",
                    oneDigitActivationEnabled ? "true" : "false"
                  );
                }
              } catch {}

              if (uiSettings) {
                await setUiSettingsInStorage(uiSettings);
              }

              if (favorites) {
                await setFavoritesInStorage(favorites);
              }

              cachedState = await getJumpToState();
            });
          })().catch((err) => console.error("cancel background handler failed:", err));
          return;
        }
      });

      reply({ type: "parentReady" });
      event.completed();
    }
  );
}


Office.actions.associate("openJumpDialog", openJumpDialog);