/*
  commands.js – Option B engine with cache + refresh-on-open (signature based)
*/

import {
  getJumpToState,
  toggleFavorite as toggleFavoriteInStorage,
  recordActivation,
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
        try { dialog.messageChild(JSON.stringify(msg)); } catch {}
      };

      const flushStateQueue = async () => {
        if (cachedState) {
          while (pendingStateRequests.length) {
            pendingStateRequests.pop();
            reply({ type: "stateData", state: cachedState });
          }
        }

        const changed = await ensureFreshState();
        if (changed && cachedState) {
          reply({ type: "stateData", state: cachedState });
        }
      };

      dialog.addEventHandler(
        Office.EventType.DialogMessageReceived,
        async (arg) => {
          let msg;
          try { msg = JSON.parse(arg.message); } catch { return; }

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
              reply({ type: "stateData", state: cachedState });
            });
            return;
          }

          if (msg.type === "selectSheet") {
            dialog.close();
            await withLock(async () => {
              await activateSheetById(msg.sheetId);
              await recordActivation(msg.sheetId);
            });
            event.completed();
            return;
          }

          if (msg.type === "cancel") {
            dialog.close();
            event.completed();
            return;
          }
        }
      );

      reply({ type: "parentReady" });
      event.completed();
    }
  );
}

Office.actions.associate("openJumpDialog", openJumpDialog);
