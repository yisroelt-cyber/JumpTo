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
  if (!baseState) return baseState;

  const activeId = await getActiveWorksheetId();

  const sheetsArr = Array.isArray(baseState.sheets) ? baseState.sheets : [];
  const visibleIds = new Set(sheetsArr.map((s) => s.id));
  const idToName = new Map(sheetsArr.map((s) => [s.id, s.name]));

  const nRaw = baseState.settings?.recentsDisplayCount;
  const n = Number.isFinite(nRaw) ? Math.max(1, Math.min(20, Math.floor(nRaw))) : 20;

  const baseRecents = Array.isArray(baseState.recents) ? baseState.recents : [];
  const recentIds = baseRecents
    .map((r) => (typeof r === "string" ? r : r?.id))
    .filter(Boolean);

  const filtered = [];
  for (const id of recentIds) {
    if (id === activeId) continue;
    if (!visibleIds.has(id)) continue;
    filtered.push(id);
    if (filtered.length >= n) break;
  }

  return {
    ...baseState,
    recents: filtered.map((id) => ({ id, name: idToName.get(id) || "" })),
  };
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

      const flushStateQueue = async () => {
        if (cachedState) {
          const state = await buildDialogState(cachedState);
          while (pendingStateRequests.length) {
            pendingStateRequests.pop();
            reply({ type: "stateData", state });
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
            await setUiSettingsInStorage(patch);
            if (!cachedState) {
              cachedState = await getJumpToState();
            } else {
              cachedState = {
                ...cachedState,
                settings: { ...(cachedState.settings || {}), ...patch },
              };
            }
            const state = await buildDialogState(cachedState);
            reply({ type: "stateData", state });
          });
          return;
        }

        if (msg.type === "selectSheet") {
          dialog.close();
          await withLock(async () => {
            if (msg.uiSettings && typeof msg.uiSettings === "object") {
              await setUiSettingsInStorage(msg.uiSettings);
              if (cachedState) {
                cachedState = {
                  ...cachedState,
                  settings: { ...(cachedState.settings || {}), ...msg.uiSettings },
                };
              }
            }
            await activateSheetById(msg.sheetId);
            await recordActivation(msg.sheetId);
          });
          event.completed();
          return;
        }

        if (msg.type === "cancel") {
          dialog.close();
          await withLock(async () => {
            if (msg.uiSettings && typeof msg.uiSettings === "object") {
              await setUiSettingsInStorage(msg.uiSettings);
              if (cachedState) {
                cachedState = {
                  ...cachedState,
                  settings: { ...(cachedState.settings || {}), ...msg.uiSettings },
                };
              }
            }
          });
          event.completed();
          return;
        }
      });

      reply({ type: "parentReady" });
      event.completed();
    }
  );
}


Office.actions.associate("openJumpDialog", openJumpDialog);