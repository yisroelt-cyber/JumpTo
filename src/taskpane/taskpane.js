/* taskpane.js: instrumentation */
if (typeof window !== "undefined") {
  window.addEventListener("error", (e) => {
    console.error("JumpTo startup error:", e?.error || e?.message || e);
  });
  window.addEventListener("unhandledrejection", (e) => {
    console.error("JumpTo unhandled rejection:", e?.reason || e);
  });
  console.log("JumpTo: taskpane.js loaded");
}

/* global Office, Excel */

let dialogRef = null;
let officeReadyPromise = null;


function setRibbonDebug(text) {
  const el = document.getElementById("ribbonDebug");
  if (el) el.textContent = text || "";
}

async function refreshRibbonDebug() {
  try {
    if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage) {
      setRibbonDebug("Ribbon debug: OfficeRuntime.storage unavailable");
      return;
    }
    const ts = await OfficeRuntime.storage.getItem("jtLastRibbonInvoke");
    const count = await OfficeRuntime.storage.getItem("jtRibbonInvokeCount");
    if (ts) {
      setRibbonDebug(`Ribbon debug: last invoke ${ts} (count ${count || "1"})`);
    } else {
      setRibbonDebug("Ribbon debug: no invoke recorded yet");
    }
  } catch {
    // ignore
  }
}

function setStatus(text) {
  const el = document.getElementById("status");
  if (el) el.textContent = text || "";
}

function setButtonState(disabled, label) {
  const btn = document.getElementById("openDialogBtn");
  if (!btn && !shimMode) return;
  btn.disabled = !!disabled;
  if (label) btn.textContent = label;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function officeReady() {
  if (typeof Office === "undefined") {
    return Promise.reject(new Error("Office.js is not available in this taskpane."));
  }
  if (!officeReadyPromise) {
    officeReadyPromise = Office.onReady();
  }
  return officeReadyPromise;
}

async function getVisibleWorksheets() {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name,items/visibility");
    await context.sync();
    return sheets.items
      .filter((ws) => ws.visibility === "Visible")
      .map((ws) => ({ name: ws.name }));
  });
}

async function activateSheetByName(sheetName) {
  return Excel.run(async (context) => {
    context.workbook.worksheets.getItem(sheetName).activate();
    await context.sync();
  });
}

function safeMessageChild(payload) {
  try {
    dialogRef?.messageChild(JSON.stringify(payload));
  } catch (e) {
    // Best-effort; dialog might be closing.
  }
}

function closeDialog() {
  try {
    dialogRef?.close();
  } catch (e) {
    // ignore
  }
  dialogRef = null;
}

function openDialogOnce() {
  return new Promise((resolve, reject) => {
    try {
      Office.context.ui.displayDialogAsync(
        `${window.location.origin}/dialog.html`,
        { height: 70, width: 45, displayInIframe: true },
        (result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            reject(result.error || new Error("Couldn't open dialog"));
            return;
          }

          dialogRef = result.value;

          dialogRef.addEventHandler(Office.EventType.DialogEventReceived, () => {
            closeDialog();
          });

          dialogRef.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
            let msg;
            try {
              msg = JSON.parse(arg.message);
            } catch {
              return;
            }

            try {
              if (msg.type === "ping") {
                safeMessageChild({ type: "parentReady" });
                return;
              }

              if (msg.type === "getSheets") {
                const sheets = await getVisibleWorksheets();
                safeMessageChild({ type: "sheetsData", sheets });
                return;
              }

              if (msg.type === "selectSheet") {
                await activateSheetByName(msg.sheetName);
                closeDialog();
                return;
              }
            } catch (err) {
              safeMessageChild({ type: "error", message: String(err?.message || err) });
            }
          });

          // Signal to the dialog that handlers are attached (prevents race conditions).
          safeMessageChild({ type: "parentReady" });

          resolve();
        }
      );
    } catch (e) {
      reject(e);
    }
  });
}

async function openDialogWithRetry() {
  // The first open after sideload/Excel restart can be sluggish.
  // We don't attempt to "fix" host behavior—just avoid a no-op click.
  const backoffMs = [0, 300, 900];
  let lastErr = null;

  await officeReady();

  for (let i = 0; i < backoffMs.length; i += 1) {
    if (backoffMs[i] > 0) {
      await sleep(backoffMs[i]);
    }

    try {
      await openDialogOnce();
      return;
    } catch (e) {
      lastErr = e;
    }
  }

  throw lastErr || new Error("Couldn't open dialog");
}

function wireUI() {
  const params = new URLSearchParams(window.location.search);
  const shimMode = params.get("shim") === "1";
  if (shimMode) {
    try {
      document.documentElement.classList.add("jt-shim");
    } catch {}
  }

  const btn = document.getElementById("openDialogBtn");
  if (!btn) return;

  if (btn) btn.addEventListener("click", () => {
    setButtonState(true, "Initializing…");
    setStatus("");

    openDialogWithRetry()
      .then(() => {
        setStatus("");
      })
      .catch((e) => {
        setStatus(`Couldn't open dialog. ${e?.message || e}`);
      })
      .finally(() => {
        setButtonState(false, "Open JumpTo");
      });
  });

  // AUTO-SHIM: when opened from the ribbon via ShowTaskpane, immediately open the modal dialog.
  if (shimMode) {
    // Defer a tick so DOM paints before we attempt to open the dialog.
    setTimeout(() => {
      openDialogWithRetry().catch(() => {});
    }, 0);
  }

}

(function init() {
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", wireUI);
  } else {
    wireUI();
  }

  // Ribbon debug polling.
  try {
    refreshRibbonDebug();
    setInterval(refreshRibbonDebug, 1000);
  } catch {}

  // Warm Office in background.
  if (typeof Office !== "undefined") {
    try {
      officeReady().catch(() => {});
    } catch {}
  }
})();
