/* commands.js: instrumentation */
if (typeof window !== "undefined") {
  window.addEventListener("error", (e) => {
    console.error("JumpTo startup error:", e?.error || e?.message || e);
  });
  window.addEventListener("unhandledrejection", (e) => {
    console.error("JumpTo unhandled rejection:", e?.reason || e);
  });
  console.log("JumpTo: commands.js loaded");
}

/* global Office, Excel, OfficeRuntime */

// Option B performance patch:
// - Open dialog immediately (no artificial delay)
// - Start preloading state concurrently with dialog creation
// - Push stateData proactively (dialog doesn't need to request)
// - Close dialog immediately after activation; recordActivation after close

import {
  getJumpToState,
  toggleFavorite as toggleFavoriteInStorage,
  recordActivation,
} from "../services/jumpToStorage";

const JT_BUILD = "39";

// --- Simple single-flight lock (VBA: "only one macro at a time") ---
let lockBusy = false;
const lockQueue = [];

function withLock(fn) {
  return new Promise((resolve, reject) => {
    lockQueue.push({ fn, resolve, reject });
    void pump();
  });
}

async function pump() {
  if (lockBusy) return;
  const job = lockQueue.shift();
  if (!job) return;

  lockBusy = true;
  try {
    const result = await job.fn();
    job.resolve(result);
  } catch (e) {
    job.reject(e);
  } finally {
    lockBusy = false;
    void pump();
  }
}

async function activateSheetById(sheetId) {
  // Excel JS does not reliably expose getItemById for worksheets.
  // Map id -> name, then activate by name.
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/id,name,visibility");
    await context.sync();

    const match = sheets.items.find((ws) => ws.id === sheetId);
    if (!match) {
      throw new Error(`Worksheet not found (id: ${sheetId})`);
    }

    context.workbook.worksheets.getItem(match.name).activate();
    await context.sync();
  });
}

function action(event) {
  try {
    console.log("JumpTo: action command executed");
  } finally {
    event.completed();
  }
}

function openJumpDialog(event) {
  // Diagnostics: confirm ribbon command handler is invoked.
  try {
    const ts = new Date().toISOString();
    console.log(`[JT][build ${JT_BUILD}] openJumpDialog invoked`, ts, window.location.href);
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      OfficeRuntime.storage.setItem("jtBuild", JT_BUILD).catch(() => {});
      OfficeRuntime.storage.setItem("jtLastRibbonInvoke", ts).catch(() => {});
      OfficeRuntime.storage
        .getItem("jtRibbonInvokeCount")
        .then((v) => {
          const n = (parseInt(v || "0", 10) || 0) + 1;
          return OfficeRuntime.storage.setItem("jtRibbonInvokeCount", String(n));
        })
        .catch(() => {});
    }
  } catch {
    /* ignore */
  }

  // Build dialog URL relative to the current page.
  const params = new URLSearchParams(window.location.search);
  const v = params.get("v");
  const dialogUrlObj = new URL("./dialog.html", window.location.href);
  if (v) dialogUrlObj.searchParams.set("v", v);
  const dialogUrl = dialogUrlObj.toString();

  const options = { height: 70, width: 45, displayInIframe: true };

  let completed = false;
  const completeOnce = () => {
    if (completed) return;
    completed = true;
    try {
      event.completed();
    } catch {
      /* ignore */
    }
  };

  // Start preloading state immediately, but do NOT block dialog opening.
  // This overlaps Excel.run with the dialog webview startup.
  const preloadStatePromise = withLock(async () => {
    try {
      return await getJumpToState();
    } catch (e) {
      console.error("Preload getJumpToState failed:", e);
      return null;
    }
  });

  const tryOpen = (attempt) => {
    Office.context.ui.displayDialogAsync(dialogUrl, options, (asyncResult) => {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        const err = asyncResult.error;
        console.error("displayDialogAsync failed:", err);

        const code = err?.code;
        const transientCodes = new Set([12002, 12006, 12007, 12009]);
        if (attempt < 5 && transientCodes.has(code)) {
          const delay = 250 + attempt * 350;
          setTimeout(() => tryOpen(attempt + 1), delay);
          return;
        }

        completeOnce();
        return;
      }

      const dialog = asyncResult.value;

      const safeReply = (obj) => {
        try {
          dialog.messageChild(JSON.stringify(obj));
        } catch {
          /* ignore */
        }
      };

      const closeDialog = () => {
        try {
          dialog.close();
        } catch {
          /* ignore */
        }
      };

      const closeAndComplete = () => {
        closeDialog();
        completeOnce();
      };

      // Immediately signal parent readiness.
      safeReply({ type: "parentReady" });

      // Proactively send stateData as soon as preload finishes.
      // This removes the extra round-trip (dialog asking for sheets).
      preloadStatePromise.then((state) => {
        if (state) safeReply({ type: "stateData", state });
      });

      const sendFreshState = async () => {
        const state = await getJumpToState();
        safeReply({ type: "stateData", state });
      };

      const onMessage = async (arg) => {
        let payload;
        try {
          payload = JSON.parse(arg.message);
        } catch {
          return;
        }

        try {
          if (payload?.type === "ping") {
            safeReply({ type: "parentReady" });
            return;
          }

          if (payload?.type === "getSheets") {
            // Prefer the preloaded state if it is ready, otherwise compute fresh.
            const preloaded = await preloadStatePromise;
            if (preloaded) {
              safeReply({ type: "stateData", state: preloaded });
            } else {
              await withLock(async () => {
                await sendFreshState();
              });
            }
            return;
          }

          if (payload?.type === "toggleFavorite" && payload.sheetId) {
            await withLock(async () => {
              await toggleFavoriteInStorage(payload.sheetId);
              await sendFreshState();
            });
            return;
          }

          if (payload?.type === "selectSheet" && payload.sheetId) {
            // Activate sheet first (user-perceived success), then close immediately.
            await withLock(async () => {
              await activateSheetById(payload.sheetId);
            });

            // Close ASAP for better perceived speed.
            closeAndComplete();

            // Record activation after close (best-effort; do not block UI).
            void withLock(async () => {
              try {
                await recordActivation(payload.sheetId);
              } catch (e) {
                console.error("recordActivation failed:", e);
              }
            });

            return;
          }

          if (payload?.type === "cancel") {
            closeAndComplete();
            return;
          }
        } catch (err) {
          console.error("DialogMessageReceived handling failed:", err);
          safeReply({ type: "error", message: String(err?.message || err) });
        }
      };

      const onEvent = (evt) => {
        console.log("DialogEventReceived:", evt);
        closeAndComplete();
      };

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, onMessage);
      dialog.addEventHandler(Office.EventType.DialogEventReceived, onEvent);

      // Ribbon command can complete now; dialog stays running.
      completeOnce();
    });
  };

  // Open immediately (no artificial delay).
  tryOpen(0);
}

Office.actions.associate("action", action);
Office.actions.associate("openJumpDialog", openJumpDialog);
