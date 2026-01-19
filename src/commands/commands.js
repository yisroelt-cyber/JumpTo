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

/* global Office, Excel */

const JT_BUILD = "37";

Office.onReady(() => {
  try {
    const ts = new Date().toISOString();
    console.log(`[JT][build 37] commands host ready`, ts, window.location.href);
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      OfficeRuntime.storage.setItem("jtBuild", JT_BUILD).catch(() => {});
      OfficeRuntime.storage.setItem("jtCommandsHostReady", ts).catch(() => {});
    }
  } catch { /* ignore */ }
});

/**
 * Optional placeholder command. Safe in Excel.
 * @param {Office.AddinCommands.Event} event
 */
function action(event) {
  try {
    // No-op; keep for compatibility if referenced anywhere.
    console.log("JumpTo: action command executed");
  } finally {
    event.completed();
  }
}

/**
 * Opens the JumpTo dialog and handles sheet selection.
 * @param {Office.AddinCommands.Event} event
 */
function openJumpDialog(event) {
  // Diagnostics: confirm the ribbon command handler is actually invoked.
  try {
    const ts = new Date().toISOString();
    console.log("[JT][build 37] openJumpDialog invoked", ts, window.location.href);
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      OfficeRuntime.storage.setItem("jtLastRibbonInvoke", ts).catch(() => {});
      OfficeRuntime.storage.getItem("jtRibbonInvokeCount").then((v) => {
        const n = (parseInt(v || "0", 10) || 0) + 1;
        return OfficeRuntime.storage.setItem("jtRibbonInvokeCount", String(n));
      }).catch(() => {});
    }
  } catch { /* ignore */ }

  // IMPORTANT: Build the dialog URL relative to the current page.
  // Using window.location.origin breaks on GitHub Pages because the add-in
  // is hosted under a repo subpath (e.g. /JumpTo/), and origin would drop it.
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
    try { event.completed(); } catch { /* ignore */ }
  };

  const tryOpen = (attempt) => {
    Office.context.ui.displayDialogAsync(dialogUrl, options, (asyncResult) => {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        const err = asyncResult.error;
        console.error("displayDialogAsync failed:", err);

        // Retry a couple of times for transient dialog launch failures.
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

      const closeAndComplete = () => {
        try { dialog.close(); } catch { /* ignore */ }
        completeOnce();
      };

      const safeReply = (obj) => {
        try { dialog.messageChild(JSON.stringify(obj)); } catch { /* ignore */ }
      };

      const getVisibleSheets = async () => {
        return Excel.run(async (context) => {
          const sheets = context.workbook.worksheets;
          sheets.load("items/name,visibility");
          await context.sync();
          return sheets.items
            .filter((ws) => ws.visibility === Excel.SheetVisibility.visible)
            .map((ws) => ({ name: ws.name }));
        });
      };

      const onMessage = async (arg) => {
        let payload;
        try { payload = JSON.parse(arg.message); } catch { return; }

        try {
          if (payload?.type === "ping") {
            safeReply({ type: "parentReady" });
            return;
          }

          if (payload?.type === "getSheets") {
            const sheets = await getVisibleSheets();
            safeReply({ type: "sheetsData", sheets });
            return;
          }

          if (payload?.type === "selectSheet" && payload.sheetName) {
            try {
              await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem(payload.sheetName);
                sheet.activate();
                await context.sync();
              });
            } catch (err) {
              console.error("Sheet activation failed:", err);
              safeReply({ type: "error", message: String(err?.message || err) });
            }
            closeAndComplete();
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

      completeOnce();
    });
  };

  // Ensure Office is ready before launching (helps stability in desktop).
  setTimeout(() => tryOpen(0), 500);
}


// Register the functions with Office.
Office.actions.associate("action", action);
Office.actions.associate("openJumpDialog", openJumpDialog);
