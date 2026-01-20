/* commands.js – Option B engine (with cancel restored) */
import {
  getJumpToState,
  toggleFavorite as toggleFavoriteInStorage,
  recordActivation,
} from "../services/jumpToStorage";

let lockBusy = false;
const lockQueue = [];
const pendingStateRequests = [];

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

async function activateSheetById(sheetId) {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/id,name");
    await context.sync();
    const ws = sheets.items.find(s => s.id === sheetId);
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
        const state = await getJumpToState();
        while (pendingStateRequests.length) {
          pendingStateRequests.pop();
          reply({ type: "stateData", state });
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
              await flushStateQueue();
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

          // RESTORED: cancel handling
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
