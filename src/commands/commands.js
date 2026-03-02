// 2026-03-02 01:00 UTC
function delayMs(ms) {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
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
  detectWorkbookReadOnly,
} from "../services/jumpToStorage";

let lockBusy = false;
const lockQueue = [];
const pendingStateRequests = [];

let cachedState = null;
let cachedSignature = "";
let lastCheckTs = 0;
const CHECK_TTL_MS = 1500;

// Read-only state: detected once at dialog open, cached for the session.
let isReadOnlyCached = null;

// Global (per-user) option keys stored in OfficeRuntime.storage
const OPT_ROW_HEIGHT = "JumpTo.Option.RowHeightPreset";
const OPT_BASELINE_ORDER = "JumpTo.Option.BaselineOrder";
const OPT_FREQUENT_ON_TOP = "JumpTo.Option.FrequentOnTop";
const OPT_FAV_PERCENT = "JumpTo.Option.FavPercentManual";
const OPT_RECENTS_DISPLAY_COUNT = "JumpTo.Option.RecentsDisplayCount";
const OPT_QUICK_RETURN = "JumpTo.Option.EnableQuickReturn";

// Legacy key (previously global) for one-digit activation; now workbook-scoped.
const OPT_ONE_DIGIT_LEGACY = "JumpTo.Option.OneDigitActivation";

// Session flag: tracks whether a jump has been made since commands.js loaded.


// Module-level cache for global ORTS settings.
// Populated on first dialog open, reused for all subsequent opens in the same
// shared-runtime session. Invalidated whenever the user changes a setting.
// Keyed values mirror the OPT_* constants above.
let ortsSettingsCache = null;

function invalidateOrtsSettingsCache() {
  ortsSettingsCache = null;
}


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

// Single Excel.run that fetches everything needed at dialog-open time:
// workbook read-only status, active sheet id, and visible sheet signature.
// Replaces three previously separate Excel.run calls (detectWorkbookReadOnly,
// computeSheetSignature, getActiveWorksheetId) with one round-trip.
async function getWorkbookSnapshot() {
  return Excel.run(async (context) => {
    context.workbook.load("readOnly");
    const activeSheet = context.workbook.worksheets.getActiveWorksheet();
    activeSheet.load("id");
    const sheets = context.workbook.worksheets;
    sheets.load("items/id,name,visibility");
    await context.sync();

    const visibleItems = sheets.items.filter((s) => s.visibility === "visible");
    const signature = visibleItems.map((s) => `${s.id}:${s.name}`).join("|");

    return {
      isReadOnly: !!context.workbook.readOnly,
      activeSheetId: activeSheet.id,
      signature,
    };
  });
}

async function ensureFreshState() {
  const now = Date.now();
  if (now - lastCheckTs < CHECK_TTL_MS) return false;
  lastCheckTs = now;

  const sig = await computeSheetSignature();
  if (sig === cachedSignature && cachedState) return false;

  cachedState = await getJumpToState({ isReadOnly: !!isReadOnlyCached });
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

async function buildDialogState(baseState, activeSheetId = null) {
  if (!baseState) return baseState;

  // Use provided activeSheetId (from getWorkbookSnapshot) to avoid an extra
  // Excel.run. Fall back to a direct lookup only when not supplied.
  const activeId = activeSheetId !== null ? activeSheetId : await getActiveWorksheetId();

  const sheetsArr = Array.isArray(baseState.sheets) ? baseState.sheets : [];
  const visibleIds = new Set(sheetsArr.map((s) => s.id));
  const idToName = new Map(sheetsArr.map((s) => [s.id, s.name]));

  const baseRecents = Array.isArray(baseState.recents) ? baseState.recents : [];
  const recentIds = baseRecents
    .map((r) => (typeof r === "string" ? r : r?.id))
    .filter(Boolean);

  // Defaults (legacy fallbacks)
  let rowHeightPreset = "Standard";
  let enableQuickReturn = true; // default ON

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
    : 5;

  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage?.getItems) {
      // Use cached values if available (populated on first open, lives for the
      // shared-runtime session, invalidated on any setting change).
      if (!ortsSettingsCache) {
        const keys = [
          OPT_ROW_HEIGHT,
          OPT_BASELINE_ORDER,
          OPT_FREQUENT_ON_TOP,
          OPT_FAV_PERCENT,
          OPT_RECENTS_DISPLAY_COUNT,
          OPT_ONE_DIGIT_LEGACY,
          OPT_QUICK_RETURN,
        ];
        // getItems returns { [key]: value } for all requested keys in one call.
        ortsSettingsCache = await OfficeRuntime.storage.getItems(keys);
      }

      const v   = ortsSettingsCache[OPT_ROW_HEIGHT];
      const bo  = ortsSettingsCache[OPT_BASELINE_ORDER];
      const fot = ortsSettingsCache[OPT_FREQUENT_ON_TOP];
      const fp  = ortsSettingsCache[OPT_FAV_PERCENT];
      const rc  = ortsSettingsCache[OPT_RECENTS_DISPLAY_COUNT];
      const od  = ortsSettingsCache[OPT_ONE_DIGIT_LEGACY];
      const qr  = ortsSettingsCache[OPT_QUICK_RETURN];

      if (v)  rowHeightPreset = String(v);
      if (bo) baselineOrder = String(bo);
      if (fot === "false") frequentOnTop = false;
      else if (fot === "true") frequentOnTop = true;
      if (fp !== null && fp !== undefined && fp !== "") favPercentManual = Number(fp);
      if (rc !== null && rc !== undefined && rc !== "") recentsDisplayCount = Number(rc);
      if (qr === "false") enableQuickReturn = false;
      else if (qr === "true") enableQuickReturn = true;

      // Legacy one-digit activation: seed from global key only if workbook has no override.
      if (baseState.settings?.oneDigitActivationEnabled === undefined) {
        if (od === "false") oneDigitActivationEnabled = false;
        else if (od === "true") oneDigitActivationEnabled = true;
      }
    }
  } catch (e) {
    // If batched getItems fails, clear cache so next open retries.
    ortsSettingsCache = null;
  }

  // Clamp recentsDisplayCount for use in filtered Recents list.
  const n = Number.isFinite(recentsDisplayCount)
    ? Math.max(1, Math.min(20, Math.floor(recentsDisplayCount)))
    : 5;

  const filtered = [];
  for (const id of recentIds) {
    if (id === activeId) continue;
    if (!visibleIds.has(id)) continue;
    filtered.push(id);
    if (filtered.length >= n) break;
  }

  // Merge frequency data into sheet objects so dialog.jsx can use s.freq directly.
  const freqById = (baseState.global?.freqById && typeof baseState.global.freqById === "object")
    ? baseState.global.freqById
    : {};
  const sheetsWithFreq = sheetsArr.map((s) => ({
    ...s,
    freq: Number(freqById[s.id] || 0),
  }));

  // DIAG: write Quick Return debug info to settings sheet column E
  try {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItemOrNullObject("_JumpToAddinSettings");
      ws.load("name");
      await ctx.sync();
      if (!ws.isNullObject) {
        const ts = new Date().toISOString();
        const r0 = recentIds[0] || "(none)";
        const r1 = recentIds[1] || "(none)";
        const match = (r0 === activeId) ? "MATCH" : "NO-MATCH";
        const r0name = idToName.get(r0) || r0;
        const r1name = idToName.get(r1) || r1;
        const activeName = idToName.get(activeId) || activeId;
        ws.getRange("E1").values = [["QR DIAG (latest at top)"]];
        // Shift existing rows down by inserting at E2
        ws.getRange("E2").insert(Excel.InsertShiftDirection.down);
        ws.getRange("E2").values = [[`${ts} | active=${activeName} | r0=${r0name} | r1=${r1name} | ${match}`]];
        await ctx.sync();
      }
    });
  } catch (e) { /* diag best-effort */ }

  return {
    ...baseState,
    activeSheetId: activeId,
    sheets: sheetsWithFreq,
    // Keep workbook settings minimal; dialog UI can still display global values (provided via `global`).
    settings: { favPercentManual, recentsDisplayCount },
    global: { oneDigitActivationEnabled, rowHeightPreset, baselineOrder, frequentOnTop, devPremium: !!(baseState.global?.devPremium), enableQuickReturn },
    recents: filtered.map((id) => ({ id, name: idToName.get(id) || "" })),
    isReadOnly: !!(baseState.isReadOnly),
    // Raw recentIds (unfiltered, unsliced) needed by Quick Return logic in dialog.
    recentIds: recentIds,

  };
}


async function setGlobalUiSettings(patch) {
  const p = patch && typeof patch === "object" ? patch : {};
  if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage?.setItem) return;
  invalidateOrtsSettingsCache();

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
  // Reset read-only detection each time the dialog opens — workbook state can change
  // (e.g. user saves a read-only copy as a new writable file).
  isReadOnlyCached = null;
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
        } catch (e) {
          // ignore
        }
      };

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
      cachedState = await getJumpToState({ preferCache: true, isReadOnly: !!isReadOnlyCached });
      if (cachedState) {
        const state = await buildDialogState(cachedState);
        while (pendingStateRequests.length) {
          pendingStateRequests.pop();
          reply({ type: "stateData", state });
        }
      }
    } catch (e) {
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
        } catch (e) {
          return;
        }

        
        try {
          if (msg && msg.type && String(msg.type).startsWith("diag")) {
          }
        } catch (e) {
          // ignore
        }

        if (msg.type === "ping") {
          // Dialog pings until it knows the parent is listening.
          reply({ type: "parentReady" });
          return;
        }

        if (msg.type === "dialogReady") {
          // Dialog has registered its parent-message handler; it's now safe to send stateData.
          await withLock(async () => {
            // Single Excel.run fetches read-only status, active sheet id, and sheet
            // signature — replacing three previously separate round-trips.
            let activeSheetId = null;
            let snapshot = null;
            try {
              snapshot = await getWorkbookSnapshot();
              isReadOnlyCached = snapshot.isReadOnly;
              activeSheetId = snapshot.activeSheetId;
            } catch {
              // Fallback: individual checks if combined snapshot fails.
              if (isReadOnlyCached === null) {
                try {
                  isReadOnlyCached = await detectWorkbookReadOnly();
                } catch {
                  isReadOnlyCached = false;
                }
              }
            }
            const readOnly = !!isReadOnlyCached;

            // Use snapshot signature to decide whether a full state refresh is needed.
            let changedAny = false;
            if (snapshot) {
              const now = Date.now();
              if (now - lastCheckTs >= CHECK_TTL_MS) {
                lastCheckTs = now;
                if (snapshot.signature !== cachedSignature || !cachedState) {
                  cachedState = await getJumpToState({ isReadOnly: readOnly });
                  cachedSignature = snapshot.signature;
                  changedAny = true;
                }
              }
            } else {
              // Snapshot failed — fall back to original ensureFreshState path.
              try {
                const changed0 = await ensureFreshState();
                changedAny = changedAny || !!changed0;
              } catch (e) {
                // ignore
              }
            }

            if (!cachedState) {
              try {
                cachedState = await getJumpToState({ preferCache: true, isReadOnly: readOnly });
              } catch (e) {
                // ignore
              }
            }
            if (!cachedState) {
              cachedState = await getJumpToState({ isReadOnly: readOnly });
            }

            let state = await buildDialogState(cachedState, activeSheetId);

            // If we still look "invalid" (common right after Excel restart), retry once after a short delay.
            if (state && state.__meta && state.__meta.dts === 0) {
              try {
                await delayMs(250);
              } catch (e) {
                // ignore
              }
              try {
                const changed1 = await ensureFreshState();
                changedAny = changedAny || !!changed1;
              } catch (e) {
                // ignore
              }
              if (cachedState) {
                state = await buildDialogState(cachedState, activeSheetId);
              }
            }

            reply({ type: "stateData", state });

            if (changedAny && cachedState) {
              try {
                const state2 = await buildDialogState(cachedState, activeSheetId);
                reply({ type: "stateData", state: state2 });
              } catch (e) {
                // ignore
              }
            }
          });
          return;
        }

        if (msg.type === "setFavorites") {
          // Silently ignored in read-only — the dialog should not allow this action when read-only,
          // but we guard here as a safety net.
          if (isReadOnlyCached) return;
          const ids = Array.isArray(msg.favorites) ? msg.favorites.filter(Boolean) : [];
          await withLock(async () => {
            await setFavoritesInStorage(ids);
            if (!cachedState) {
              cachedState = await getJumpToState({ isReadOnly: false });
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
            cachedState = await getJumpToState({ isReadOnly: !!isReadOnlyCached });
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
        await OfficeRuntime.storage.setItem(OPT_ROW_HEIGHT, preset);
        invalidateOrtsSettingsCache();
      }
    } catch (e) {
          // ignore
        }
    cachedState = await getJumpToState({ isReadOnly: !!isReadOnlyCached });
    const state = await buildDialogState(cachedState);
    reply({ type: "stateData", state });
  });
  return;
}

        if (msg.type === "setOneDigitActivation") {
          // Workbook-scoped setting — silently ignored in read-only (UI control will be disabled).
          if (isReadOnlyCached) return;
          const enabled = !!msg.enabled;
          await withLock(async () => {
            // Workbook-scoped: persist in workbook settings blob.
            await setUiSettingsInStorage({ oneDigitActivationEnabled: enabled });
            if (!cachedState) {
              cachedState = await getJumpToState({ isReadOnly: false });
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

if (msg.type === "setEnableQuickReturn") {
          const enabled = msg.enabled !== false; // default true
          await withLock(async () => {
            try {
              if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage?.setItem) {
                await OfficeRuntime.storage.setItem(OPT_QUICK_RETURN, enabled ? "true" : "false");
                invalidateOrtsSettingsCache();
              }
            } catch (e) {
              // ignore
            }
            cachedState = await getJumpToState({ isReadOnly: !!isReadOnlyCached });
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
          const rowHeightDirty = !!snapshot.rowHeightDirty;

          // Close + complete immediately so the dialog feels instant.
          try {
            dialog.close();
          } catch (e) {
          // ignore
        }
          event.completed();

          // Optimistically update cachedState with the destination sheet at recentIds[0].
          // This ensures Quick Return is available immediately if the user reopens the dialog
          // before the background recordActivation write completes.
          // The background task will replace this with authoritative state from the workbook.
          if (sheetId && cachedState) {
            try {
              const prevRecents = Array.isArray(cachedState.recents) ? cachedState.recents : [];
              const prevRecentIds = prevRecents.map(r => (typeof r === "string" ? r : r?.id)).filter(Boolean);
              const nextRecentIds = [sheetId, ...prevRecentIds.filter(id => id !== sheetId)].slice(0, 20);
              const idToName = new Map((Array.isArray(cachedState.sheets) ? cachedState.sheets : []).map(s => [s.id, s.name]));
              cachedState = {
                ...cachedState,
                recents: nextRecentIds.map(id => ({ id, name: idToName.get(id) || "" })),
              };
            } catch (e) {
              cachedState = null;
            }
          } else {
            cachedState = null;
          }

          // Continue work in the background so UI close is not blocked by Excel writes.
          (async () => {
            await withLock(async () => {
              let finalRecentIds = null;
              if (sheetId) {
                // Capture the origin sheet (currently active) before jumping away.
                let originSheetId = null;
                try {
                  originSheetId = await Excel.run(async (context) => {
                    const ws = context.workbook.worksheets.getActiveWorksheet();
                    ws.load("id");
                    await context.sync();
                    return ws.id;
                  });
                } catch (e) {
                  // ignore — origin capture is best-effort
                }

                await activateSheetById(sheetId);

                // Skip recording activations in read-only workbooks — all write paths would throw.
                if (!isReadOnlyCached) {
                  // Record origin first, then destination — so destination lands at position 0
                  // (most recent), origin at position 1. Skip origin if same as destination.
                  if (originSheetId && originSheetId !== sheetId) {
                    await recordActivation(originSheetId);
                  }
                  const recResult = await recordActivation(sheetId);
                  // Capture the authoritative recentIds returned by recordActivation so we can
                  // patch cachedState without doing a full getJumpToState read afterward.
                  if (recResult && Array.isArray(recResult.recents)) {
                    finalRecentIds = recResult.recents;
                  }
                }
              }

              // Persist latest state AFTER activation so persistence work doesn't delay the jump.
              // Only persist RowHeightPreset from the snapshot if the user actually changed it
              // during this dialog session. Otherwise, a default UI value could overwrite the
              // existing global value (e.g., when the dialog closes quickly).
              if (rowHeightDirty && rowHeightPreset) {
                if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage?.setItem) {
                  await OfficeRuntime.storage.setItem(OPT_ROW_HEIGHT, rowHeightPreset);
                  invalidateOrtsSettingsCache();
                }
              }

              const oneDigitActivationEnabled = !!snapshot.oneDigitActivationEnabled;

              try {
                if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage?.setItem) {
                  await OfficeRuntime.storage.setItem(
                    "JumpTo.Option.OneDigitActivation",
                    oneDigitActivationEnabled ? "true" : "false"
                  );
                  invalidateOrtsSettingsCache();
                }
              } catch (e) {
          // ignore
        }

              if (uiSettings && !isReadOnlyCached) {
                await setUiSettingsInStorage(uiSettings);
              }

              if (favorites && !isReadOnlyCached) {
                await setFavoritesInStorage(favorites);
              }

              // Keep cache coherent for the next dialog open.
              // Prefer patching recents directly from recordActivation's return value to avoid
              // a race where getJumpToState reads the workbook before the write has landed.
              if (finalRecentIds !== null && cachedState) {
                const idToName = new Map((Array.isArray(cachedState.sheets) ? cachedState.sheets : []).map(s => [s.id, s.name]));
                cachedState = {
                  ...cachedState,
                  recents: finalRecentIds.map(id => ({ id, name: idToName.get(id) || "" })),
                };
              } else {
                cachedState = await getJumpToState({ isReadOnly: !!isReadOnlyCached });
              }
            });
})().catch((err) => console.error("selectSheet background handler failed:", err));

          return;
        }

        if (msg.type === "cancel") {
          try {
            const p = msg.payload || {};
            if (p && p.settingsSnap) {
            }
          } catch (e) {
            // ignore
          }

          const snapshot = msg.snapshot && typeof msg.snapshot === "object" ? msg.snapshot : {};
          const uiSettings = snapshot.uiSettings && typeof snapshot.uiSettings === "object" ? snapshot.uiSettings : null;
          const favorites = Array.isArray(snapshot.favorites) ? snapshot.favorites.filter(Boolean) : null;
          const rowHeightPreset = typeof snapshot.rowHeightPreset === "string" ? snapshot.rowHeightPreset : "";
          const rowHeightDirty = !!snapshot.rowHeightDirty;

          try {
            dialog.close();
          } catch (e) {
          // ignore
        }
          event.completed();

          (async () => {
            await withLock(async () => {
              if (rowHeightDirty && rowHeightPreset) {
                if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage?.setItem) {
                  await OfficeRuntime.storage.setItem(OPT_ROW_HEIGHT, rowHeightPreset);
                  invalidateOrtsSettingsCache();
                }
              }

              const oneDigitActivationEnabled = !!snapshot.oneDigitActivationEnabled;

              try {
                if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage?.setItem) {
                  await OfficeRuntime.storage.setItem(
                    "JumpTo.Option.OneDigitActivation",
                    oneDigitActivationEnabled ? "true" : "false"
                  );
                  invalidateOrtsSettingsCache();
                }
              } catch (e) {
          // ignore
        }

              if (uiSettings && !isReadOnlyCached) {
                await setUiSettingsInStorage(uiSettings);
              }

              if (favorites && !isReadOnlyCached) {
                await setFavoritesInStorage(favorites);
              }

              cachedState = await getJumpToState({ isReadOnly: !!isReadOnlyCached });
            });

                        try {
            } catch (e) {
            }
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
