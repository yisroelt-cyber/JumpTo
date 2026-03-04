// 2026-03-04 20:45 UTC
import React, { useEffect, useMemo, useRef, useState } from "react";
import { MAX_RECENTS, PREMIUM_FREQ_BUMP } from "../shared/constants";

import { createRoot } from "react-dom/client";

function canMessageParentLocal() {
  try {
    return !!(Office && Office.context && Office.context.ui && Office.context.ui.messageParent);
  } catch (e) {
    return false;
  }
}

function buildSettingsSnapForParent(globalOptions, uiFavPercentManual, uiRecentsDisplayCount, refs) {
  try {
    return {
      globalOptions: {
        rowHeightPreset: globalOptions?.rowHeightPreset,
        oneDigitActivationEnabled: globalOptions?.oneDigitActivationEnabled,
        baselineOrder: globalOptions?.baselineOrder,
        frequentOnTop: globalOptions?.frequentOnTop,
      },
      ui: {
        favPercentManual: uiFavPercentManual,
        recentsDisplayCount: uiRecentsDisplayCount,
      },
      flags: {
        globalDirty: !!refs?.globalOptionsDirtyRef?.current,
        uiDirty: !!refs?.uiSettingsDirtyRef?.current,
        rowHeightDirty: !!refs?.rowHeightDirtyRef?.current,
        prefsHydrated: !!refs?.prefsHydratedRef?.current,
        prefsHydratedFromValid: !!refs?.prefsHydratedFromValidRef?.current,
      },
    };
  } catch (e) {
    return { error: "snapshotFailed" };
  }
}


function fireAndForget(promise, tag) {
  try {
    if (promise && typeof promise.then === "function") {
      promise.catch((e) => {
        try {
          console.error('[RowHeightTrace] dialog fireAndForget error', tag || '', e);
        } catch (e2) {
      // ignore
    }
      });
    }
  } catch (e) {
    try {
      console.error('[RowHeightTrace] dialog fireAndForget failure', tag || '', e);
    } catch (e2) {
      // ignore
    }
  }
}


/* global Office */

const ROW_HEIGHT_PRESETS = {
  Compact: {
    fontSize: 10,
    lineHeight: 15,
    paddingY: 1,
    estRowHeight: 17, // 15 + 1 + 1
  },
  Standard: {
    fontSize: 12,
    lineHeight: 16,
    paddingY: 2,
    estRowHeight: 20, // legacy/current
  },
  Comfortable: {
    fontSize: 14,
    lineHeight: 20,
    paddingY: 3,
    estRowHeight: 26,
  },
  Expanded: {
    fontSize: 16,
    lineHeight: 24,
    paddingY: 4,
    estRowHeight: 32,
  },
};


function safeJsonParse(str) {
  try {
    return JSON.parse(str);
  } catch (e) {
    return null;
  }
}

function TabButton({ label, active, onClick, disabled, disabledTitle }) {
  return (
    <button
      type="button"
      onClick={disabled ? undefined : onClick}
      title={disabled ? disabledTitle : undefined}
      disabled={disabled}
      style={{
        appearance: "none",
        background: "transparent",
        border: "none",
        padding: "8px 12px",
        margin: 0,
        cursor: disabled ? "default" : "pointer",
        fontFamily: "Segoe UI, Arial, sans-serif",
        fontSize: 13,
        fontWeight: active ? 600 : 400,
        color: disabled ? "rgba(0,0,0,0.35)" : "#111",
        borderBottom: active ? "2px solid #0078d4" : "2px solid transparent",
        opacity: disabled ? 0.5 : 1,
      }}
    >
      {label}
    </button>
  );
}

function clampNumber(n, min, max) {
  const v = Number(n);
  if (!Number.isFinite(v)) return min;
  return Math.min(max, Math.max(min, v));
}

// Favorites bounce diagnostics removed; keep no-op logger to avoid runtime crashes.
function favDbgLog() { /* no-op */ }

function sameFavoriteIds(a, b) {
  if (a === b) return true;
  const aa = Array.isArray(a) ? a : [];
  const bb = Array.isArray(b) ? b : [];
  if (aa.length !== bb.length) return false;
  for (let i = 0; i < aa.length; i++) {
    const ida = aa[i]?.id;
    const idb = bb[i]?.id;
    if (ida !== idb) return false;
  }
  return true;
}

function DialogApp() {
  const receivedStateDataRef = useRef(false);
  const devPremiumRef = useRef(false);

  // Suppress hover highlight on dialog open so a row under the cursor at open
  // time doesn't steal faux-focus. Cleared after layout settles (same approach
  // as XLL's Dispatcher.BeginInvoke(DispatcherPriority.Input)).
  const suppressHoverRef = useRef(true);
  useEffect(() => {
    requestAnimationFrame(() => { suppressHoverRef.current = false; });
  }, []);

  useEffect(() => {
    if (!canMessageParentLocal()) return;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);


  // =====================
  // Settings trace probe (diagnostics-only, hard-guaranteed)
  // Ensures we can verify dialog is able to write SettingsTraceLog in ORTS.
  // =====================
  const [allSheets, setAllSheets] = useState([]);
  const [activeSheetId, setActiveSheetId] = useState(null);
  const [favorites, setFavorites] = useState([]);
  // Favorites bounce fix: prevent stale parent-state hydration from overwriting recent UI edits.
  const favoritesDirtyRef = useRef(false);
  const lastUiFavMutationAtRef = useRef(0);

  const [recents, setRecents] = useState([]);
  const [recentIds, setRecentIds] = useState([]); // raw unfiltered IDs from parent
  const [globalOptions, setGlobalOptions] = useState({ oneDigitActivationEnabled: true, rowHeightPreset: "Standard", baselineOrder: "workbook", frequentOnTop: true });
  const [enableQuickReturn, setEnableQuickReturn] = useState(true);

  const [query, setQuery] = useState("");
  const [status, setStatus] = useState("Loading…");
  const [isActivating, setIsActivating] = useState(false);
  const [initError, setInitError] = useState("");
  const [activeTab, setActiveTab] = useState("Navigation");
  const [isReadOnly, setIsReadOnly] = useState(false);
  const [readOnlyBannerDismissed, setReadOnlyBannerDismissed] = useState(false);
  
  // Favorites tab UI state (remember selection across tab switches)
  const [favTabSelectedAvailableId, setFavTabSelectedAvailableId] = useState(null);
  const [favTabSelectedFavoriteId, setFavTabSelectedFavoriteId] = useState(null);

  // Hover highlight state (Navigation + Favorites tab)
  const [hoverNavFavoriteId, setHoverNavFavoriteId] = useState(null);
  const [hoverNavRecentId, setHoverNavRecentId] = useState(null);
  const [hoverFavTabAvailableId, setHoverFavTabAvailableId] = useState(null);
  const [hoverFavTabFavoriteId, setHoverFavTabFavoriteId] = useState(null);

  // UI layout settings (Navigation + Favorites tab right column)
  const [uiFavPercentManual, setUiFavPercentManual] = useState(50); // 20..80 (Favorites share when space is limited)
  const [uiRecentsDisplayCount, setUiRecentsDisplayCount] = useState(5); // 1..MAX_RECENTS
  const uiSettingsPersistTimerRef = useRef(null);

  // Global options persistence (debounced): rowHeightPreset.
  const globalOptionsPersistTimerRef = useRef(null);

  // Measured layout: keep dialog from scrolling; listboxes scroll internally
  const rootRef = useRef(null);
  const tabsRef = useRef(null);
  const footerRef = useRef(null);
  const bodyRef = useRef(null);
  const favTabFavListRef = useRef(null);
  const favTabPendingScrollIdRef = useRef(null);
  const [panelHeight, setPanelHeight] = useState(320); // computed at runtime


  // Favorites persistence (Favorites tab): debounce writes to minimize sheet churn
  const favPersistTimerRef = useRef(null);
  const favDirtyRef = useRef(false);
  const favoritesRef = useRef([]);

  // Faux-focus: keyboard focus stays on search box at all times, but a virtual
  // focus cycles between the three Navigation tab lists via Tab.
  // "all" = All Sheets, "favorites" = Favorites, "recents" = Recents
  const [fauxFocus, setFauxFocus] = useState("all");
  // Independent highlight indices — each section remembers its position independently.
  const [highlightAll, setHighlightAll] = useState(0);
  const [highlightFav, setHighlightFav] = useState(0);
  const [highlightRec, setHighlightRec] = useState(0);
  // Legacy alias kept for Favorites tab (unchanged).
  const highlightIndex = highlightAll;
  const setHighlightIndex = setHighlightAll;
  // Scroll container refs for right-column lists (used for scrollIntoView on arrow nav).
  const navFavListRef = useRef(null);
  const navRecListRef = useRef(null);
  // Row element refs for right-column lists.
  const navFavRowRefs = useRef([]);
  const navRecRowRefs = useRef([]);
  const requestedRef = useRef(false);
  const timeoutIdRef = useRef(null);
  const statusRef = useRef("Loading…");
  const sheetsLenRef = useRef(0);
  const searchInputRef = useRef(null);
  const listRowRefs = useRef([]);
  const focusTimersRef = useRef([]);
    const parentReadyRef = useRef(false);
  const uiSettingsReadyRef = useRef(false);
  const uiSettingsDirtyRef = useRef(false);
  const uiSettingsDirtyDesiredRef = useRef(null);
  const globalOptionsDirtyRef = useRef(false);
  const globalOptionsDirtyDesiredRef = useRef(null);
  const rowHeightDirtyRef = useRef(false);
  const prefsHydratedRef = useRef(false);
  const prefsHydratedFromValidRef = useRef(false);
  // Tracks hydration of ORTS-backed per-user UI prefs (fav slider + recents count).
  // These are safe to apply even when workbook meta validity is false.
  const uiPrefsHydratedRef = useRef(false);
  
  const refsForParentSnap = { globalOptionsDirtyRef, uiSettingsDirtyRef, rowHeightDirtyRef, prefsHydratedRef, prefsHydratedFromValidRef, uiPrefsHydratedRef };
useEffect(() => { favoritesRef.current = favorites; }, [favorites]);

  useEffect(() => { statusRef.current = status; }, [status]);

  // settings change watcher (throttled via Column C flush cadence; ORTS append is bounded)
  useEffect(() => {
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [globalOptions?.rowHeightPreset, globalOptions?.oneDigitActivationEnabled, globalOptions?.baselineOrder, globalOptions?.frequentOnTop, uiFavPercentManual, uiRecentsDisplayCount]);


  // settings change watcher (throttled)
  useEffect(() => {
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [globalOptions?.rowHeightPreset, globalOptions?.oneDigitActivationEnabled, globalOptions?.baselineOrder, globalOptions?.frequentOnTop, uiFavPercentManual, uiRecentsDisplayCount]);

  useEffect(() => { sheetsLenRef.current = allSheets.length; }, [allSheets]);

  const requestSearchFocus = (reason = "") => {
    // Office dialog webviews can be finicky with focus timing. Be defensive and never throw.
    // Cancel any existing scheduled focus attempts.
    try {
      (focusTimersRef.current || []).forEach((t) => window.clearTimeout(t));
    } catch (e) {
      // ignore
    }
    focusTimersRef.current = [];

    const tryFocus = () => {
      const el = searchInputRef.current;
      if (!el || typeof el.focus !== "function") return;
      try {
        el.focus();
      } catch (e) {
        // ignore
      }
    };

    // Immediate + short delayed retries (escalating).
    tryFocus();
    const delays = [50, 120, 250, 450, 750, 1100];
    delays.forEach((ms) => {
      const t = window.setTimeout(tryFocus, ms);
      focusTimersRef.current.push(t);
    });

    if (reason) {
      // Useful breadcrumb for troubleshooting focus timing in Office webviews.
      // Keep as a debug log only; does not affect UX.
      // eslint-disable-next-line no-console
      console.debug("[JumpToSheet][Dialog] requestSearchFocus:", reason);
    }
  };

  // Minimal crash visibility: surface unexpected issues in the console and (optionally) in the dialog.
  useEffect(() => {
    const onError = (evt) => {
      try {
        const msg = evt?.message || "Unknown error";
        console.error("[JumpToSheet][Dialog] window.onerror:", msg, evt);
        setInitError((prev) => prev || msg);
      } catch (e) {
        // ignore
      }
    };
    const onUnhandled = (evt) => {
      try {
        const reason = evt?.reason;
        const msg = reason?.message || String(reason || "Unhandled promise rejection");
        console.error("[JumpToSheet][Dialog] unhandledrejection:", msg, evt);
        setInitError((prev) => prev || msg);
      } catch (e) {
        // ignore
      }
    };

    window.addEventListener("error", onError);
    window.addEventListener("unhandledrejection", onUnhandled);

    return () => {
      window.removeEventListener("error", onError);
      window.removeEventListener("unhandledrejection", onUnhandled);
      try {
        (focusTimersRef.current || []).forEach((t) => window.clearTimeout(t));
      } catch (e) {
        // ignore
      }
      focusTimersRef.current = [];
    };
  }, []);

  // Compute panel height so the dialog itself never scrolls (controls scroll internally).
  useEffect(() => {
    const compute = () => {
      try {
        const body = bodyRef.current;
        if (!body) return;
        const bodyRect = body.getBoundingClientRect();
        const h = Math.max(220, Math.floor(bodyRect.height));
        setPanelHeight(h);
      } catch (e) {
        // ignore
      }
    };
    compute();
    const onResize = () => compute();
    window.addEventListener("resize", onResize);
    let ro = null;
    try {
      if (window.ResizeObserver) {
        ro = new ResizeObserver(() => compute());
        if (bodyRef.current) ro.observe(bodyRef.current);
      }
    } catch (e) {
      // ignore
    }
    return () => {
      window.removeEventListener("resize", onResize);
      try {
        if (ro) ro.disconnect();
      } catch (e) {
      // ignore
    }
    };
  }, []);


  // Office dialog webviews sometimes ignore the HTML autoFocus attribute.
  // Use a small focus retry sequence to reliably place the caret in the search box.
  useEffect(() => {
    if (activeTab !== "Navigation") return;
    requestSearchFocus("mount");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);


  useEffect(() => {
    let disposed = false;
    let pingTimer = null;
    let pingCount = 0;

    const canMessageParent = () => {
      try {
        return !!(
          window.Office &&
          Office.context &&
          Office.context.ui &&
          typeof Office.context.ui.messageParent === "function"
        );
      } catch (e) {
        return false;
      }
    };

    const requestSheets = () => {
      // Only attempt to talk to the parent after Office is actually initialized.
      if (!canMessageParent()) {
        requestedRef.current = false;
        setStatus("Initializing Office…");
        window.setTimeout(requestSheets, 100);
        return;
      }
      if (requestedRef.current) return;
      requestedRef.current = true;
      try {
        Office.context.ui.messageParent(JSON.stringify({ type: "getSheets" }));
      } catch (err) {
        console.error("messageParent(getSheets) failed:", err);
        setStatus("Unable to contact parent.");
      }
    };


function snapshotDialogSettings(globalOptions, uiFavPercentManual, uiRecentsDisplayCount, flags) {
  try {
    return {
      globalOptions: {
        rowHeightPreset: globalOptions?.rowHeightPreset,
        oneDigitActivationEnabled: globalOptions?.oneDigitActivationEnabled,
        baselineOrder: globalOptions?.baselineOrder,
        frequentOnTop: globalOptions?.frequentOnTop,
      },
      ui: {
        favPercentManual: uiFavPercentManual,
        recentsDisplayCount: uiRecentsDisplayCount,
      },
      flags: flags || {},
    };
  } catch (e) {
    return { error: "snapshotFailed" };
  }
}


    const sendPing = () => {
      if (!canMessageParent()) return;
      try {
        Office.context.ui.messageParent(JSON.stringify({ type: "ping" }));
      } catch (e) {
        // ignore
      }
    };

    if (window.Office && typeof Office.onReady === "function") {
      Office.onReady(() => {
      try {
      } catch (e) { /* ignore */ }
      if (disposed) return;

      // Listen for parent responses.
      Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        (arg) => {
          if (disposed) return;
          const msg = safeJsonParse(arg?.message);
          if (!msg?.type) return;

          if (msg.type === "parentReady") {
            parentReadyRef.current = true;
            if (pingTimer) {
              window.clearInterval(pingTimer);
              pingTimer = null;
            }
            requestSheets();
            // Re-assert focus after the parent handshake.
            if (activeTab === "Navigation" || activeTab === "Favorites") requestSearchFocus("parentReady");
            return;
          }

          if (msg.type === "stateData") {
            try {
              receivedStateDataRef.current = true;
              devPremiumRef.current = !!(msg.state?.global?.devPremium);
            } catch (e) {
              // ignore
            }

            const state = msg.state || {};
            // Update read-only flag whenever parent sends state.
            if (typeof state.isReadOnly === "boolean") setIsReadOnly(state.isReadOnly);
            // Update active sheet ID
            if (state.activeSheetId) setActiveSheetId(state.activeSheetId);
            // Update Quick Return session state

            const sheets = Array.isArray(state.sheets) ? state.sheets : [];
            setAllSheets(sheets);
            setFavorites((prev) => {
            const next = Array.isArray(state.favorites) ? state.favorites : [];
            if (favoritesDirtyRef.current && !sameFavoriteIds(prev, next)) {
              const ageMs = Date.now() - (lastUiFavMutationAtRef.current || 0);
              if (ageMs < 2000) {
                // We very recently changed favorites locally; ignore stale parent hydration that would "bounce" the UI.
                favDbgLog("hydrate:parentState:skippedDirty", prev, next);
                return prev;
              }
              // If we're still dirty after a while, accept parent as authoritative to avoid getting "stuck" forever.
              favoritesDirtyRef.current = false;
            }
            // If parent has caught up to our local favorites, clear dirty.
            if (favoritesDirtyRef.current && sameFavoriteIds(prev, next)) {
              favoritesDirtyRef.current = false;
            }
            favDbgLog("hydrate:parentState", prev, next);
            return next;
          });
            setRecents(Array.isArray(state.recents) ? state.recents : []);
            // Raw recentIds for Quick Return logic (unfiltered, unsliced)
            if (Array.isArray(state.recentIds)) setRecentIds(state.recentIds);
            const incomingMeta = state && typeof state === "object" ? state.__meta : null;
            const incomingSettingsValid = !!incomingMeta?.settingsValid;
            const incomingFavoritesValid = !!incomingMeta?.favoritesValid;

            const hasIncomingGlobal = !!(state && state.global && typeof state.global === "object");
            const hasIncomingUiPrefs = !!(state && state.settings && typeof state.settings === "object");

            // Layered hydration:
            // A) ORTS-backed globals: safe even when workbook meta validity is false.
            // B) Workbook/UI-dependent sections: only when meta validity passes.
            const hydrateGlobal = hasIncomingGlobal && !prefsHydratedRef.current;
            const hydrateWorkbookSections = incomingSettingsValid || incomingFavoritesValid;

            if (hydrateGlobal) {
              prefsHydratedRef.current = true;
            }
            if (hydrateWorkbookSections && incomingSettingsValid) {
              prefsHydratedFromValidRef.current = true;
            }


            setGlobalOptions((prev) => {
              const incoming = state.global || { oneDigitActivationEnabled: true, rowHeightPreset: "Standard" };
              // Also update enableQuickReturn from global (not inside globalOptions, kept separate).
              if (typeof incoming.enableQuickReturn === "boolean") {
                setEnableQuickReturn(incoming.enableQuickReturn);
              }
              // If the user has changed global options locally (e.g. clicked a checkbox) and we're still waiting
              // for parent persistence to catch up, don't let late-arriving stateData overwrite the user's intent.
                            if (globalOptionsDirtyRef.current) {
                const desired = globalOptionsDirtyDesiredRef.current;
                if (
                  desired &&
                  !!incoming.oneDigitActivationEnabled === !!desired.oneDigitActivationEnabled &&
                  String(incoming.rowHeightPreset || "Standard") === String(desired.rowHeightPreset || "Standard")
                ) {
                  // Parent has caught up; accept incoming and clear dirty.
                  globalOptionsDirtyRef.current = false;
                  globalOptionsDirtyDesiredRef.current = null;
                  rowHeightDirtyRef.current = false;
                  return incoming;
                }
                // Ignore stale incoming globals while dirty.
                return prev || incoming;
              }
              return incoming;
            });

            // ORTS-backed per-user UI preferences (fav slider + recents count)
            // These are safe to apply even when workbook meta validity is false.
            const hydrateUiPrefs = hasIncomingUiPrefs && !uiPrefsHydratedRef.current;
            if (hydrateUiPrefs) {
              uiPrefsHydratedRef.current = true;
            }

            // UI settings (persisted per-user)
            if (hydrateUiPrefs) {
            try {
              const ui = state.settings || {};              const favPct = Number.isFinite(Number(ui.favPercentManual)) ? Number(ui.favPercentManual) : 50;
              const recCnt = Number.isFinite(Number(ui.recentsDisplayCount)) ? Number(ui.recentsDisplayCount) : 10;
              const incomingFav = Math.min(80, Math.max(20, Math.round(favPct)));
              const incomingCnt = Math.min(MAX_RECENTS, Math.max(1, Math.round(recCnt)));

const incomingBaseOrder = (ui.baselineOrder === "alpha" ? "alpha" : "workbook");
const incomingFrequentOnTop = !!ui.frequentOnTop;

              if (uiSettingsDirtyRef.current) {
                const desired = uiSettingsDirtyDesiredRef.current;
                if (
                  desired &&
                  Math.min(80, Math.max(20, Math.round(desired.favPercentManual))) === incomingFav &&
                  Math.min(MAX_RECENTS, Math.max(1, Math.round(desired.recentsDisplayCount))) === incomingCnt
                ) {
                  uiSettingsDirtyRef.current = false;
                  uiSettingsDirtyDesiredRef.current = null;
                  setUiFavPercentManual(incomingFav);
                  setUiRecentsDisplayCount(incomingCnt);
                } else {
                  // Ignore stale incoming UI settings while dirty.
                }
              } else {
                setUiFavPercentManual(incomingFav);
                setUiRecentsDisplayCount(incomingCnt);
              }

              setGlobalOptions((prev) => ({
                ...(prev || {}),
                baselineOrder: incomingBaseOrder,
                frequentOnTop: incomingFrequentOnTop,
              }));
            } catch (e) {
              // ignore
            }

            }
            uiSettingsReadyRef.current = uiPrefsHydratedRef.current || prefsHydratedFromValidRef.current || uiSettingsDirtyRef.current;
            setStatus(sheets.length ? "" : "No visible worksheets found.");

            // Re-assert focus after data arrives (this is the moment users start typing).
            if (activeTab === "Navigation" || activeTab === "Favorites") requestSearchFocus("sheetsData");
            if (timeoutIdRef.current) {
              window.clearTimeout(timeoutIdRef.current);
              timeoutIdRef.current = null;
            }
            if (pingTimer) {
              window.clearInterval(pingTimer);
              pingTimer = null;
            }
            return;
          }

          if (msg.type === "error") {
            setIsActivating(false);
            setStatus(msg.message || "An error occurred.");
            return;
          }
        },
        () => {
          try {
            if (canMessageParent()) {
              Office.context.ui.messageParent(JSON.stringify({ type: "dialogReady" }));
            }
          } catch (e) {
            // ignore
          }
        }
      );

      // Ping until parent is ready (prevents races where parent hasn't attached message handlers yet).
      sendPing();
      pingTimer = window.setInterval(() => {
        if (disposed) return;
        if (parentReadyRef.current) return;
        pingCount += 1;
        sendPing();
        if (pingCount >= 25) { // ~10s
          window.clearInterval(pingTimer);
          pingTimer = null;
          if (statusRef.current === "Loading…" && sheetsLenRef.current === 0) {
            setStatus(
              "Still loading… If this doesn’t resolve, close this dialog and launch it again from the ribbon command (Home → JumpTo)."
            );
          }
        }
      }, 400);

      // Defensive timeout, but only if we never get a response.
      timeoutIdRef.current = window.setTimeout(() => {
        if (disposed) return;
        if (statusRef.current === "Loading…" && sheetsLenRef.current === 0) {
          setStatus(
            "Still loading… If this doesn’t resolve, close this dialog and launch it again from the ribbon command (Home → JumpTo)."
          );
        }
      }, 12000);
    });
    } else {
      // Office.js may not be loaded yet in some dialog webviews (race with script loading).
      // We'll retry initialization shortly rather than rendering a broken UI.
      window.setTimeout(() => {
        try { requestSheets(); } catch (e) {
      // ignore
    }
      }, 100);
    }

    return () => {
      disposed = true;
      if (pingTimer) window.clearInterval(pingTimer);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

    const computeTier = (freq) => {
    const f = Number(freq || 0);
    if (f < 10) return 0;
    return 1 + Math.floor(Math.log(f / 10) / Math.log(1.35));
  };

  const filtered = useMemo(() => {
    const q = (query || "").toLowerCase();
    let items = Array.isArray(allSheets) ? [...allSheets] : [];

    // Exclude active sheet unless it's the only visible sheet and search is empty
    const isOnlySheet = items.length === 1;
    const searchActive = !!q;
    if (activeSheetId && (!isOnlySheet || searchActive)) {
      items = items.filter((s) => s?.id !== activeSheetId);
    }

    if (q) {
      items = items.filter((s) => (s?.name || "").toLowerCase().includes(q));
    }

    // Base order: workbook order (default) or alphabetical
    const baselineOrder = String(globalOptions?.baselineOrder || "workbook");
    if (baselineOrder === "alpha") {
      items.sort((a, b) => (a?.name || "").localeCompare(b?.name || ""));
    } else {
      items.sort((a, b) => Number(a?.orderIndex || 0) - Number(b?.orderIndex || 0));
    }

    // Apply "frequent bump" ONLY when search is active, the list is narrowed, AND premium is enabled.
    // Premium is active if the compile-time constant is set OR the dev magic cell (A10 = "DEV_PREMIUM") is present.
    const allCount = Array.isArray(allSheets) ? allSheets.length : 0;
    if (q && items.length < allCount && (PREMIUM_FREQ_BUMP || devPremiumRef.current) && !!(globalOptions?.frequentOnTop)) {
      const N = items.length;
      const k = Math.min(Math.max(Math.ceil(0.1 * N), 1), 5); // candidates considered; does not force a bump

      // Base order index for stable tie-breaks
      const idxById = new Map(items.map((s, i) => [s?.id, i]));

      // Top-k candidates by frequency
      const byFreq = items.slice().sort((a, b) => Number(b?.freq || 0) - Number(a?.freq || 0));
      const candidates = byFreq.slice(0, k);
      const others = byFreq.slice(k);

      const medianOf = (arr) => {
        const nums = arr
          .map((s) => Number(s?.freq || 0))
          .filter((n) => Number.isFinite(n))
          .sort((a, b) => a - b);
        if (!nums.length) return 0;
        const mid = Math.floor(nums.length / 2);
        return nums.length % 2 ? nums[mid] : (nums[mid - 1] + nums[mid]) / 2;
      };

      const baselineMed = others.length ? medianOf(others) : medianOf(byFreq);

      // Dynamic-ish ratio: stricter for tiny N, relaxes toward ~1.75 for larger lists.
      const t = Math.max(0, Math.min(1, (N - 2) / 18));
      const ratio = 2.5 - 0.75 * t;

      const threshold = Math.max(5, baselineMed * ratio);

      let bumped = candidates.filter((s) => Number(s?.freq || 0) >= threshold);

      // Sort bumped by frequency desc, then base order; cap to 5.
      bumped.sort((a, b) => {
        const df = Number(b?.freq || 0) - Number(a?.freq || 0);
        if (df !== 0) return df;
        return (idxById.get(a?.id) ?? 0) - (idxById.get(b?.id) ?? 0);
      });
      if (bumped.length > 5) bumped = bumped.slice(0, 5);

      if (bumped.length) {
        const bumpedIds = new Set(bumped.map((s) => s?.id));
        const rest = items.filter((s) => !bumpedIds.has(s?.id));
        items = [...bumped, ...rest];
      }
    }

    // Quick Return: prepend a special row when conditions are met (search must be empty).
    // Quick Return: enabled, search empty, recentIds[0] == activeSheetId (still on jumped-to sheet), recentIds[1] exists and visible.
    if (
      enableQuickReturn &&
      !q &&
      Array.isArray(recentIds) &&
      recentIds.length >= 2 &&
      recentIds[0] === activeSheetId
    ) {
      const returnId = recentIds[1];
      const returnSheet = items.find((s) => s?.id === returnId);
      if (returnSheet) {
        // Remove from its normal position so it appears exactly once.
        items = items.filter((s) => s?.id !== returnId);
        // Prepend Quick Return row.
        items = [
          { ...returnSheet, name: returnSheet.name + "  ↩", isQuickReturn: true },
          ...items,
        ];
      }
    }

    return items;
  }, [allSheets, query, globalOptions?.baselineOrder, activeSheetId, enableQuickReturn, recentIds]);

  const favoriteIds = useMemo(() => new Set((favorites || []).map((f) => f?.id).filter(Boolean)), [favorites]);

  // Right column sizing controls (Favorites/Recents split)
  const favPercentEffective = Math.min(80, Math.max(20, Math.round(uiFavPercentManual)));
  const recPercentEffective = 100 - favPercentEffective;

  // Row height metrics (applies to all listboxes).
  const activePresetName = String(globalOptions?.rowHeightPreset || "Standard");
  const activeRowPreset = ROW_HEIGHT_PRESETS[activePresetName] || ROW_HEIGHT_PRESETS.Standard;
  const rowFontSize = activeRowPreset.fontSize;
  const rowLineHeight = activeRowPreset.lineHeight;
  const rowPadY = activeRowPreset.paddingY;
  const rowEstHeightPx = activeRowPreset.estRowHeight;

  // Layout constants (px) – tuned for Office dialog webviews
  const LABEL_ROW_H = 18;
  const GAP_H = 6;
  const NAV_MID_GAP_H = 10; // extra breathing room between Favorites list and Recents label (Nav tab right column)

  const ROW_EST_H = rowEstHeightPx; // estimated row height for a single list item (padding + lineHeight + border)

  // Favorites tab right column height budget.
// We split the right column into 70% favorites list + 30% controls.
// Important: the "Favorites" label + margins consume extra vertical space,
// so subtract a small overhead from the panelHeight to avoid clipping.
const favTabRightOverhead = (LABEL_ROW_H * 1) + (GAP_H * 2);
const favTabListsTotal = Math.max(140, Math.floor(panelHeight - favTabRightOverhead));

// Favorites tab right column:
// - Top: Favorites list (scrolls internally)
// - Bottom: Controls block (Up/Down + transfer guidance)
//
// Layout rule (current): fixed 70/30 split.
// Rationale: give the Favorites listbox most of the real estate; keep controls anchored low.
const favTabFavListHeight = Math.max(80, Math.floor(favTabListsTotal * 0.70));
const favTabBottomBlockHeight = Math.max(80, favTabListsTotal - favTabFavListHeight);

  // Navigation tab right column: two scenarios
  //  1) No-conflict: show all (subject to minimum shares), ignore ratio/settings; put any extra space in the middle.
  //  2) Conflict: apply user-selected policy (fixed ratio with surplus-donation, or prioritize Favorites up to 80%).
  const navRightOverhead = (LABEL_ROW_H * 2) + (GAP_H * 3);
  const navRightH = Math.max(140, Math.floor(panelHeight - navRightOverhead));

  const navFavMin = Math.max(60, Math.floor(navRightH * 0.20));
  const navRecMin = Math.max(60, Math.floor(navRightH * 0.20));

  const navFavRowsNeed = Math.max(1, (Array.isArray(favorites) ? favorites : []).length);
  const navRecRowsNeed = Math.max(
    1,
    Math.min((Array.isArray(recents) ? recents : []).length, uiRecentsDisplayCount)
  );

  const navFavNeed = Math.max(navFavMin, (navFavRowsNeed * ROW_EST_H) + 8);
  const navRecNeed = Math.max(navRecMin, (navRecRowsNeed * ROW_EST_H) + 8);

  let navTabHasExtraSpace = false;
  let navTabFavListHeight = navFavMin;
  let navTabRecListHeight = navRecMin;

  if (navFavNeed + navRecNeed <= navRightH) {
    // No conflict – show all and push the extra into the middle spacer.
    navTabHasExtraSpace = true;
    navTabFavListHeight = navFavNeed;
    navTabRecListHeight = navRecNeed;
  } else {
    // Conflict – apply fixed ratio (20..80 ↔ 80..20), with "surplus donation" (do not waste rows on the side that does not need them).
    navTabHasExtraSpace = false;

      // Fixed ratio (20..80 ↔ 80..20), with "surplus donation" (do not waste rows on the side that doesn't need them).
      let favH = Math.floor((navRightH * favPercentEffective) / 100);
      let recH = navRightH - favH;

      // Enforce minimums.
      if (favH < navFavMin) { favH = navFavMin; recH = navRightH - favH; }
      if (recH < navRecMin) { recH = navRecMin; favH = navRightH - recH; }

      // Donate surplus only (never take below what the other side needs).
      if (navRecNeed < recH) {
        const surplus = recH - navRecNeed;
        // Give surplus to Favorites, but only if Favorites needs it.
        if (navFavNeed > favH) {
          const give = Math.min(surplus, navFavNeed - favH);
          favH += give;
          recH = navRightH - favH;
        }
      } else if (navFavNeed < favH) {
        const surplus = favH - navFavNeed;
        if (navRecNeed > recH) {
          const give = Math.min(surplus, navRecNeed - recH);
          recH += give;
          favH = navRightH - recH;
        }
      }

      // Final safety clamps.
      if (favH < navFavMin) { favH = navFavMin; recH = navRightH - favH; }
      if (recH < navRecMin) { recH = navRecMin; favH = navRightH - recH; }

      navTabFavListHeight = favH;
      navTabRecListHeight = recH;
  }


  const isFavorite = (sheetId) => favoriteIds.has(sheetId);

  const addFavoriteLocal = (sheetId) => {
    favoritesDirtyRef.current = true;
    lastUiFavMutationAtRef.current = Date.now();

    setFavorites((prev) => {
      const arr = Array.isArray(prev) ? prev : [];
      let next = arr;

      if (arr.some((x) => x?.id === sheetId)) {
        next = arr;
      } else {
        const s = (Array.isArray(allSheets) ? allSheets : []).find((x) => x?.id === sheetId);
        const name = s?.name || "";
        next = [...arr, { id: sheetId, name }];
      }

      favDbgLog("ui:addFavorite", prev, next);
      return next;
    });

    setFavTabSelectedFavoriteId(sheetId);
    setFavTabSelectedAvailableId(null);
    favTabPendingScrollIdRef.current = sheetId;
    schedulePersistFavorites("add");
  };

  const removeFavoriteLocal = (sheetId) => {
    favoritesDirtyRef.current = true;
    lastUiFavMutationAtRef.current = Date.now();

    setFavorites((prev) => {
      const next = (Array.isArray(prev) ? prev : []).filter((x) => x?.id !== sheetId);
      favDbgLog("ui:removeFavorite", prev, next);
      return next;
    });

    if (favTabSelectedFavoriteId === sheetId) setFavTabSelectedFavoriteId(null);
    schedulePersistFavorites("remove");
  };

  const moveFavoriteLocal = (sheetId, direction) => {
    favoritesDirtyRef.current = true;
    lastUiFavMutationAtRef.current = Date.now();
    if (direction !== "up" && direction !== "down") return;

    setFavorites((prev) => {
      const arr = Array.isArray(prev) ? prev.slice() : [];
      let next = arr;

      const idx = arr.findIndex((x) => x?.id === sheetId);
      if (idx < 0) {
        next = arr;
      } else {
        const to = direction === "up" ? idx - 1 : idx + 1;
        if (to < 0 || to >= arr.length) {
          next = arr;
        } else {
          const [item] = arr.splice(idx, 1);
          arr.splice(to, 0, item);
          next = arr;
        }
      }

      favDbgLog(`ui:moveFavorite:${direction}`, prev, next);
      return next;
    });

    schedulePersistFavorites("move");
  };

  const sendSetFavoritesToParent = (ids) => {
    try {
      Office.context.ui.messageParent(JSON.stringify({ type: "setFavorites", favorites: ids }));
    } catch (err) {
      console.error("messageParent(setFavorites) failed:", err);
    }
  };

  const sendSetUiSettingsToParent = (settings) => {
    try {
      Office.context.ui.messageParent(JSON.stringify({ type: "setUiSettings", settings }));
    } catch (err) {
      console.error("messageParent(setUiSettings) failed:", err);
    }
  };

  const schedulePersistUiSettings = (reason) => {
    if (!uiSettingsReadyRef.current && !uiSettingsDirtyRef.current) return;
    if (uiSettingsPersistTimerRef.current) {
      clearTimeout(uiSettingsPersistTimerRef.current);
    }
    uiSettingsPersistTimerRef.current = setTimeout(() => {
      uiSettingsPersistTimerRef.current = null;
      try {
        sendSetUiSettingsToParent({
          favPercentManual: Math.min(80, Math.max(20, Math.round(uiFavPercentManual))),
          recentsDisplayCount: Math.min(MAX_RECENTS, Math.max(1, Math.round(uiRecentsDisplayCount))),
          baselineOrder: (globalOptions?.baselineOrder === "alpha" ? "alpha" : "workbook"),
          frequentOnTop: !!(globalOptions?.frequentOnTop),
        });
      } catch (e) {
        // ignore
      }
    }, 700);
  };

  const flushPersistUiSettingsNow = (reason) => {
    if (uiSettingsPersistTimerRef.current) {
      clearTimeout(uiSettingsPersistTimerRef.current);
      uiSettingsPersistTimerRef.current = null;
    }

    try {
      sendSetUiSettingsToParent({
        favPercentManual: Math.min(80, Math.max(20, Math.round(uiFavPercentManual))),
        recentsDisplayCount: Math.min(MAX_RECENTS, Math.max(1, Math.round(uiRecentsDisplayCount))),
        baselineOrder: (globalOptions?.baselineOrder === "alpha" ? "alpha" : "workbook"),
        frequentOnTop: !!(globalOptions?.frequentOnTop),
      });
    } catch (e) {
      // ignore
    }
  };


  const schedulePersistGlobalOptions = (reason) => {
    if (!globalOptionsDirtyRef.current) return;
    if (globalOptionsPersistTimerRef.current) {
      clearTimeout(globalOptionsPersistTimerRef.current);
    }
    globalOptionsPersistTimerRef.current = setTimeout(() => {
      globalOptionsPersistTimerRef.current = null;
      try {
        const preset = String(globalOptions?.rowHeightPreset || "Standard");

        try {

          if (Office?.context?.ui?.messageParent) {
            Office.context.ui.messageParent(JSON.stringify({ type: "setRowHeightPreset", preset, __src: "dialog:schedule:" + String(reason || "") }));
            Office.context.ui.messageParent(JSON.stringify({ type: "setOneDigitActivation", enabled: !!(globalOptions?.oneDigitActivationEnabled) }));

          }

        } catch (err) {

          console.error("messageParent(setRowHeightPreset) failed:", err);

        }
      } catch (e) {
        // ignore
      }
    }, 600);
  };

  const flushPersistGlobalOptionsNow = (reason) => {
    if (globalOptionsPersistTimerRef.current) {
      clearTimeout(globalOptionsPersistTimerRef.current);
      globalOptionsPersistTimerRef.current = null;
    }
    try {
      const preset = String(globalOptions?.rowHeightPreset || "Standard");

      try {

        if (Office?.context?.ui?.messageParent) {

          Office.context.ui.messageParent(JSON.stringify({ type: "setRowHeightPreset", preset, __src: "dialog:flush:" + String(reason || "") }));
          Office.context.ui.messageParent(JSON.stringify({ type: "setOneDigitActivation", enabled: !!(globalOptions?.oneDigitActivationEnabled) }));

        }

      } catch (err) {

        console.error("messageParent(setRowHeightPreset) failed:", err);

      }} catch (e) {
      // ignore
    }
  };

  // Lock the dialog viewport: prevent the browser (body/html) from scrolling.
  // Office dialogs can have the default browser body margin, which creates a
  // small page scroll that hides the search box. Per the LPD, the dialog frame
  // must not scroll in Navigation/Favorites; only internal listboxes may scroll.
  useEffect(() => {
    try {
      const html = document.documentElement;
      const body = document.body;
      if (html) {
        html.style.height = "100%";
        html.style.overflow = "hidden";
      }
      if (body) {
        body.style.margin = "0";
        body.style.height = "100%";
        body.style.overflow = "hidden";
      }
    } catch (e) {
      // ignore
    }
  }, []);

  // Persist UI settings when they change (debounced).
  useEffect(() => {
    if (!parentReadyRef.current) return;
    schedulePersistUiSettings("ui-change");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [uiFavPercentManual, uiRecentsDisplayCount, globalOptions?.baselineOrder, globalOptions?.frequentOnTop]);

  // Persist global options when they change (debounced).
  useEffect(() => {
    if (!parentReadyRef.current) return;
    schedulePersistGlobalOptions("global-change");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [globalOptions?.rowHeightPreset, globalOptions?.oneDigitActivationEnabled]);

  // Expose flush for Save & Close
  useEffect(() => {
    window.flushPersistUiSettingsNow = flushPersistUiSettingsNow;
    return () => { try { delete window.flushPersistUiSettingsNow; } catch (e) {
      // ignore
    } };
  }, [uiFavPercentManual, uiRecentsDisplayCount, globalOptions?.baselineOrder, globalOptions?.frequentOnTop]);

  useEffect(() => {
    window.flushPersistGlobalOptionsNow = flushPersistGlobalOptionsNow;
    return () => { try { delete window.flushPersistGlobalOptionsNow; } catch (e) {
      // ignore
    } };
  }, [globalOptions?.rowHeightPreset, globalOptions?.oneDigitActivationEnabled]);

  // Favorites tab: when a new favorite is added, keep it selected and scroll it into view.
  useEffect(() => {
    try {
      if (activeTab !== "Favorites") return;
      const id = favTabPendingScrollIdRef.current;
      if (!id) return;

      // Defer until after layout/paint so the row exists.
      const doScroll = () => {
        const host = favTabFavListRef.current;
        if (!host) return false;
        const el = host.querySelector(`[data-sheetid="${String(id)}"]`);
        if (!el) return false;
        try {
          el.scrollIntoView({ block: "nearest" });
        } catch (e) {
          // ignore
        }
        return true;
      };

      // Try immediately, then on next frame if needed.
      if (doScroll()) {
        favTabPendingScrollIdRef.current = null;
        return;
      }
      const raf = window.requestAnimationFrame(() => {
        if (doScroll()) favTabPendingScrollIdRef.current = null;
      });
      return () => window.cancelAnimationFrame(raf);
    } catch (e) {
      // ignore
    }
  }, [activeTab, favorites]);


  const schedulePersistFavorites = (reason) => {
    favDirtyRef.current = true;
    if (favPersistTimerRef.current) {
      clearTimeout(favPersistTimerRef.current);
    }
    favPersistTimerRef.current = setTimeout(() => {
      favPersistTimerRef.current = null;
      try {
        const items = (Array.isArray(favoritesRef.current) ? favoritesRef.current : []).filter(x => x?.id);
        sendSetFavoritesToParent(items);
        favDirtyRef.current = false;
      } catch (e) {
        // ignore
      }
    }, 900);
  };

  const flushPersistFavoritesNow = (reason) => {
    if (!favDirtyRef.current) return;
    if (favPersistTimerRef.current) {
      clearTimeout(favPersistTimerRef.current);
      favPersistTimerRef.current = null;
    }
    try {
      const items = (Array.isArray(favoritesRef.current) ? favoritesRef.current : []).filter(x => x?.id);
      sendSetFavoritesToParent(items);
    } catch (e) {
      // ignore
    }
    favDirtyRef.current = false;
  };

  const rowStyle = {
    padding: `${rowPadY}px 10px`,
    fontSize: rowFontSize,
    lineHeight: `${rowLineHeight}px`,
    cursor: isActivating ? "default" : "pointer",
    borderBottom: "1px solid rgba(0,0,0,0.06)",
    userSelect: "none",
    opacity: isActivating ? 0.65 : 1,
  };


  

  // Build a "last known state" snapshot for actions that may close the dialog quickly
  // (e.g., selecting a sheet or cancelling). This avoids needing to flush multiple debounced
  // persistence paths before the action can proceed.
  const buildPersistSnapshot = () => {
    const uiSettings = {
      favPercentManual: Math.min(80, Math.max(20, Math.round(uiFavPercentManual))),
      recentsDisplayCount: Math.min(MAX_RECENTS, Math.max(1, Math.round(uiRecentsDisplayCount))),
      baselineOrder: (globalOptions?.baselineOrder === "alpha" ? "alpha" : "workbook"),
      frequentOnTop: !!(globalOptions?.frequentOnTop),
    };

    const favoritesItems = (Array.isArray(favoritesRef.current) ? favoritesRef.current : [])
      .filter((x) => x?.id);

    const rowHeightPreset = String(globalOptions?.rowHeightPreset || "Standard");
    const oneDigitActivationEnabled = !!(globalOptions?.oneDigitActivationEnabled);

    const rowHeightDirty = !!rowHeightDirtyRef.current;

    return { uiSettings, favorites: favoritesItems, rowHeightPreset, rowHeightDirty, oneDigitActivationEnabled };
  };

const onSelect = (sheet) => {
  if (!sheet || isActivating) return;
  const sheetId = typeof sheet === "string" ? sheet : sheet.id;
  if (!sheetId) return;

  setIsActivating(true);
  setStatus("Loading sheet…");

  try {
    const snapshot = buildPersistSnapshot();
    Office.context.ui.messageParent(JSON.stringify({ type: "selectSheet", sheetId, snapshot }));
  } catch (err) {
    console.error("messageParent(selectSheet) failed:", err);
    setIsActivating(false);
    setStatus("Failed to activate sheet.");
  }
};

const onToggleFavorite = (sheetId) => {
  if (!sheetId) return;
  try {
    Office.context.ui.messageParent(JSON.stringify({ type: "toggleFavorite", sheetId }));
  } catch (err) {
    console.error("messageParent(toggleFavorite) failed:", err);
  }
};

const onCancel = () => {
  try {
    const snapshot = buildPersistSnapshot();
    Office.context.ui.messageParent(JSON.stringify({ type: "cancel", snapshot, payload: { settingsSnap: (() => { const s = buildSettingsSnapForParent(globalOptions, uiFavPercentManual, uiRecentsDisplayCount, refsForParentSnap); try { s.flags = s.flags || {}; s.flags.receivedStateData = !!receivedStateDataRef.current; } catch (e) { /* ignore */ } return s; })(), note: { receivedStateData: !!receivedStateDataRef.current } } }));
  } catch (e) {
    // ignore
  }
};

  // When faux-focus moves to a section, clamp that section's index to valid bounds.
  useEffect(() => {
    if (activeTab !== "Navigation") return;
    if (fauxFocus === "favorites") {
      const max = Math.max(0, (Array.isArray(favorites) ? favorites : []).length - 1);
      setHighlightFav((prev) => Math.min(prev, max));
    } else if (fauxFocus === "recents") {
      const max = Math.max(0, (Array.isArray(recents) ? recents : []).slice(0, uiRecentsDisplayCount).length - 1);
      setHighlightRec((prev) => Math.min(prev, max));
    } else if (fauxFocus === "all") {
      const max = Math.max(0, (filtered?.length || 0) - 1);
      setHighlightAll((prev) => Math.min(prev, max));
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fauxFocus, activeTab]);

  // Reset faux-focus to All Sheets when switching to Navigation tab.
  useEffect(() => {
    if (activeTab === "Navigation") setFauxFocus("all");
  }, [activeTab]);

  // Reset All Sheets highlight to top when the filtered list changes (query changed).
  // Favorites and Recents positions are intentionally left alone.
useEffect(() => {
  if (activeTab !== "Navigation") return;
  setHighlightAll(0);
  requestSearchFocus("resetHighlight");
  // eslint-disable-next-line react-hooks/exhaustive-deps
}, [filtered.length, activeTab]);

// Scroll the highlighted All Sheets row into view on arrow nav.
useEffect(() => {
  if (activeTab !== "Navigation" || fauxFocus !== "all") return;
  const el = listRowRefs.current?.[highlightAll];
  if (el && typeof el.scrollIntoView === "function") {
    try { el.scrollIntoView({ block: "nearest" }); } catch (e) { /* ignore */ }
  }
}, [highlightAll, activeTab, fauxFocus]);

// Scroll highlighted Favorites row into view on arrow nav.
useEffect(() => {
  if (activeTab !== "Navigation" || fauxFocus !== "favorites") return;
  const el = navFavRowRefs.current?.[highlightFav];
  if (el && typeof el.scrollIntoView === "function") {
    try { el.scrollIntoView({ block: "nearest" }); } catch (e) { /* ignore */ }
  }
}, [highlightFav, activeTab, fauxFocus]);

// Scroll highlighted Recents row into view on arrow nav.
useEffect(() => {
  if (activeTab !== "Navigation" || fauxFocus !== "recents") return;
  const el = navRecRowRefs.current?.[highlightRec];
  if (el && typeof el.scrollIntoView === "function") {
    try { el.scrollIntoView({ block: "nearest" }); } catch (e) { /* ignore */ }
  }
}, [highlightRec, activeTab, fauxFocus]);

return (
    <div ref={rootRef} style={{ fontFamily: "Segoe UI, Arial, sans-serif", padding: 14, height: "100vh", boxSizing: "border-box", overflow: "hidden", display: "flex", flexDirection: "column" }}>
      {!!initError && (
        <div
          style={{
            marginBottom: 10,
            padding: "8px 10px",
            borderRadius: 6,
            border: "1px solid rgba(180, 0, 0, 0.35)",
            background: "rgba(255, 0, 0, 0.06)",
            fontSize: 12,
            lineHeight: 1.35,
          }}
        >
          <div style={{ fontWeight: 600, marginBottom: 4 }}>Dialog error</div>
          <div style={{ opacity: 0.9, wordBreak: "break-word" }}>{initError}</div>
          <div style={{ marginTop: 6 }}>
            <button
              type="button"
              onClick={() => setInitError("")}
              style={{
                fontSize: 12,
                padding: "4px 8px",
                borderRadius: 6,
                border: "1px solid rgba(0,0,0,0.15)",
                background: "white",
                cursor: "pointer",
              }}
            >
              Dismiss
            </button>
          </div>
        </div>
      )}

      {isReadOnly && !readOnlyBannerDismissed && (
        <div
          style={{
            marginBottom: 10,
            padding: "8px 10px",
            borderRadius: 6,
            border: "1px solid rgba(180, 130, 0, 0.35)",
            background: "rgba(255, 200, 0, 0.08)",
            fontSize: 12,
            lineHeight: 1.45,
            display: "flex",
            alignItems: "flex-start",
            gap: 8,
          }}
        >
          <div style={{ flex: "1 1 auto" }}>
            <div style={{ fontWeight: 600, marginBottom: 3 }}>This workbook is read-only</div>
            <div style={{ opacity: 0.85 }}>JumpTo is running in degraded mode. Navigation works normally, but favorites cannot be edited.</div>
          </div>
          <button
            type="button"
            onClick={() => setReadOnlyBannerDismissed(true)}
            title="Dismiss"
            style={{
              flexShrink: 0,
              fontSize: 14,
              lineHeight: 1,
              padding: "2px 4px",
              border: "none",
              background: "transparent",
              cursor: "pointer",
              color: "rgba(0,0,0,0.45)",
              marginTop: 1,
            }}
          >
            ×
          </button>
        </div>
      )}

      <div ref={tabsRef}
        style={{
          display: "flex",
          alignItems: "center",
          borderBottom: "1px solid rgba(0,0,0,0.15)",
          marginBottom: 10,
          marginTop: 2,
        }}
        role="tablist"
        aria-label="JumpTo tabs"
      >
        <TabButton label="Navigation" active={activeTab === "Navigation"} onClick={() => setActiveTab("Navigation")} />
        <TabButton
          label="Favorites"
          active={activeTab === "Favorites"}
          onClick={() => setActiveTab("Favorites")}
          disabled={isReadOnly}
          disabledTitle="Favorites cannot be edited in a read-only workbook"
        />
        <TabButton label="Settings" active={activeTab === "Settings"} onClick={() => setActiveTab("Settings")} />
      </div>

      <div ref={bodyRef} style={{ flex: "1 1 auto", overflow: "hidden" }}>

      {activeTab === "Navigation" && (
        <>
          <div style={{ display: "flex", gap: 16, height: panelHeight, overflowX: "auto", overflowY: "hidden" }}>
            {/* Left: Search + All results */}
            <div style={{ flex: "1 1 44%", minWidth: 240, paddingRight: 16, borderRight: "1px solid #d0d0d0", display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
              <div style={{ marginBottom: 10 }}>
                <input
                  autoFocus
                  ref={searchInputRef}
                  value={query}
                  onChange={(e) => {
                    setQuery(e.target.value);
                    // Typing always snaps faux-focus back to All Sheets.
                    setFauxFocus("all");
                  }}
                  onBlur={() => requestSearchFocus("fav-search-blur")}
                  onKeyDown={(e) => {
                    try {
                      const key = e.key;
                      const hasFavs = (Array.isArray(favorites) ? favorites : []).length > 0;
                      const recentsVisible = (Array.isArray(recents) ? recents : []).slice(0, uiRecentsDisplayCount);
                      const hasRecs = recentsVisible.length > 0;

                      // Tab / Shift+Tab cycles faux-focus forward / backward: all -> favorites (if any) -> recents (if any) -> all
                      if (key === "Tab") {
                        e.preventDefault();
                        if (e.shiftKey) {
                          // Backward
                          setFauxFocus((prev) => {
                            if (prev === "all") {
                              if (hasRecs) return "recents";
                              if (hasFavs) return "favorites";
                              return "all";
                            }
                            if (prev === "recents") {
                              if (hasFavs) return "favorites";
                              return "all";
                            }
                            // favorites -> all
                            return "all";
                          });
                        } else {
                          // Forward
                          setFauxFocus((prev) => {
                            if (prev === "all") {
                              if (hasFavs) return "favorites";
                              if (hasRecs) return "recents";
                              return "all";
                            }
                            if (prev === "favorites") {
                              if (hasRecs) return "recents";
                              return "all";
                            }
                            // recents -> all
                            return "all";
                          });
                        }
                        requestSearchFocus("tab");
                        return;
                      }

                      if (key === "ArrowDown") {
                        e.preventDefault();
                        if (fauxFocus === "all") {
                          setHighlightAll((prev) => Math.min(Math.max(0, (filtered?.length || 0) - 1), prev + 1));
                        } else if (fauxFocus === "favorites") {
                          setHighlightFav((prev) => Math.min(Math.max(0, (Array.isArray(favorites) ? favorites : []).length - 1), prev + 1));
                        } else if (fauxFocus === "recents") {
                          setHighlightRec((prev) => Math.min(Math.max(0, recentsVisible.length - 1), prev + 1));
                        }
                        return;
                      }

                      if (key === "ArrowUp") {
                        e.preventDefault();
                        if (fauxFocus === "all") {
                          setHighlightAll((prev) => Math.max(0, prev - 1));
                        } else if (fauxFocus === "favorites") {
                          setHighlightFav((prev) => Math.max(0, prev - 1));
                        } else if (fauxFocus === "recents") {
                          setHighlightRec((prev) => Math.max(0, prev - 1));
                        }
                        return;
                      }

                      if (key === "Enter") {
                        e.preventDefault();
                        if (fauxFocus === "all") {
                          const idx = Math.max(0, Math.min((filtered?.length || 1) - 1, highlightAll));
                          const s = filtered?.[idx];
                          if (s) onSelect(s);
                        } else if (fauxFocus === "favorites") {
                          const favList = Array.isArray(favorites) ? favorites : [];
                          const f = favList[Math.max(0, Math.min(favList.length - 1, highlightFav))];
                          if (f?.id) onSelect(f);
                        } else if (fauxFocus === "recents") {
                          const r = recentsVisible[Math.max(0, Math.min(recentsVisible.length - 1, highlightRec))];
                          if (r?.id) onSelect(r);
                        }
                        return;
                      }

                      const mods = e.altKey || e.ctrlKey || e.metaKey;
                      const oneDigit = globalOptions?.oneDigitActivationEnabled;
                      const q = query || "";
                      const leadingSpace = q.startsWith(" ");

                      // One-digit activation: always available regardless of faux-focus.
                      // Supports both main keyboard digits and numpad (Numpad0–Numpad9).
                      if (oneDigit && !mods && !leadingSpace && q === "") {
                        let digitChar = null;
                        if (key >= "0" && key <= "9") {
                          digitChar = key;
                        } else if (key.startsWith("Numpad") && key.length === 7) {
                          const c = key[6];
                          if (c >= "0" && c <= "9") digitChar = c;
                        }
                        if (digitChar !== null) {
                          const idx = digitChar === "0" ? 9 : (Number(digitChar) - 1);
                          const fav = favorites?.[idx];
                          if (fav?.id) {
                            e.preventDefault();
                            onSelect(fav);
                            return;
                          }
                        }
                      }

                      if (key === "Escape") {
                        e.preventDefault();
                        if ((query || "") !== "") {
                          setQuery("");
                        } else {
                          onCancel();
                        }
                      }
                    } catch (e) {
                      // ignore
                    }
                  }}
                  placeholder="Search sheets…"
                  disabled={!!status && status !== "" && allSheets.length === 0}
                  style={{
                    width: "100%",
                    padding: "6px 8px",
                    fontSize: 12,
                    boxSizing: "border-box",
                  }}
                />
              </div>

              {!!initError && (
                <div
                  style={{
                    marginBottom: 10,
                    padding: "8px 10px",
                    background: "rgba(232, 17, 35, 0.08)",
                    border: "1px solid rgba(232, 17, 35, 0.25)",
                    borderRadius: 6,
                    color: "#a80000",
                    fontSize: 12,
                  }}
                >
                  {initError}
                </div>
              )}

              {!!status && status !== "" ? (
                <div
                  style={{
                    padding: "10px 12px",
                    border: "1px solid rgba(0,0,0,0.1)",
                    borderRadius: 6,
                    fontSize: 13,
                    opacity: 0.9,
                  }}
                >
                  {status}
                </div>
              ) : (
                <div
                  style={{
                    flex: "1 1 auto",
                    minHeight: 0,
                    overflowY: "auto",
                    overflowX: "hidden",
                    overscrollBehavior: "contain",
                    border: fauxFocus === "all" ? "1px solid rgba(0,120,212,0.55)" : "1px solid rgba(0,0,0,0.1)",
                    borderRadius: 6,
                    boxShadow: fauxFocus === "all" ? "0 0 0 2px rgba(0,120,212,0.12)" : "none",
                  }}
                >
                  {filtered.map((s, i) => (
                    <div
                      key={s.id || s.name}
                      ref={(el) => { listRowRefs.current[i] = el; }}
                      onMouseEnter={() => { if (suppressHoverRef.current) return; try { setFauxFocus("all"); setHighlightAll(i); } catch (e) { /* ignore */ } }}
                      onClick={() => { if (!isActivating) { try { setFauxFocus("all"); setHighlightAll(i); } catch (e) { /* ignore */ } onSelect(s); } }}
                      style={{
                        ...rowStyle,
                        background: fauxFocus === "all" && i === highlightAll ? "rgba(0,120,212,0.12)" : "transparent",
                        fontStyle: s.isQuickReturn ? "italic" : "normal",
                      }}
                      role="button"
                      tabIndex={0}
                      onKeyDown={(e) => {
                        if (isActivating) return;
                        if (e.key === "Enter" || e.key === " ") onSelect(s);
                      }}
                    >
                      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <div title={s.name} style={{ flex: "1 1 auto", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{s.name}</div>
                      </div>
                    </div>
                  ))}
                  {filtered.length === 0 && (
                    <div style={{ padding: "10px 12px", fontSize: 13, opacity: 0.8 }}>
                      No matches.
                    </div>
                  )}
                </div>
              )}

            </div>

            {/* Right: Favorites + Recents */}
            <div style={{ flex: "0 0 45%", minWidth: 220, height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>

              <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 6, opacity: 0.85 }}>Favorites</div>
              <div
                ref={navFavListRef}
                style={{
                  flex: "0 1 auto",
                  height: navTabFavListHeight,
                  maxHeight: navTabFavListHeight,
                  minHeight: 0,
                  overscrollBehavior: "contain",
                  overflowY: "auto",
                  overflowX: "hidden",
                  boxSizing: "border-box",
                  border: fauxFocus === "favorites" ? "1px solid rgba(0,120,212,0.55)" : "1px solid rgba(0,0,0,0.1)",
                  borderRadius: 6,
                  marginBottom: 6,
                  boxShadow: fauxFocus === "favorites" ? "0 0 0 2px rgba(0,120,212,0.12)" : "none",
                }}>
                {(Array.isArray(favorites) ? favorites : []).map((f, i) => {
                  const slot = i < 9 ? String(i + 1) : i === 9 ? "0" : "-";
                  const name = f?.name || "";
                  const id = f?.id;
                  const isFauxHighlighted = fauxFocus === "favorites" && i === highlightFav;
                  return (
                    <div
                      key={id || `${name}_${i}`}
                      ref={(el) => { navFavRowRefs.current[i] = el; }}
                      onClick={() => { if (!isActivating && id) { setFauxFocus("favorites"); setHighlightFav(i); onSelect({ id }); } }}
                      onMouseEnter={() => { if (suppressHoverRef.current) return; setFauxFocus("favorites"); setHighlightFav(i); setHoverNavFavoriteId(id); }}
                      onMouseLeave={() => setHoverNavFavoriteId(null)}
                      style={{ ...rowStyle, background: isFauxHighlighted ? "rgba(0,120,212,0.12)" : (hoverNavFavoriteId === id && fauxFocus !== "favorites" ? "rgba(0,120,212,0.07)" : "transparent") }}
                      role="button"
                      tabIndex={0}
                      onKeyDown={(e) => {
                        if (isActivating) return;
                        if (e.key === "Enter" || e.key === " ") id && onSelect({ id });
                      }}
                    >
                      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <div style={{ width: 18, opacity: 0.75, textAlign: "right" }}>{slot}</div>
                        <div title={name} style={{ flex: "1 1 auto", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{name}</div>
                      </div>
                    </div>
                  );
                })}
                {(Array.isArray(favorites) ? favorites : []).length === 0 && (
                  <div style={{ padding: "10px 12px", fontSize: 13, opacity: 0.75 }}>No favorites yet.</div>
                )}
              </div>

              <div style={{ flex: navTabHasExtraSpace ? "1 1 auto" : `0 0 ${NAV_MID_GAP_H}px`, minHeight: NAV_MID_GAP_H }} />

              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
                <div style={{ fontSize: 12, fontWeight: 600, opacity: 0.85 }}>Recents</div>
              </div>
              <div
                ref={navRecListRef}
                style={{
                  flex: "0 1 auto",
                  height: navTabRecListHeight,
                  maxHeight: navTabRecListHeight,
                  minHeight: 0,
                  overscrollBehavior: "contain",
                  overflowY: "auto",
                  overflowX: "hidden",
                  boxSizing: "border-box",
                  border: fauxFocus === "recents" ? "1px solid rgba(0,120,212,0.55)" : "1px solid rgba(0,0,0,0.1)",
                  borderRadius: 6,
                  boxShadow: fauxFocus === "recents" ? "0 0 0 2px rgba(0,120,212,0.12)" : "none",
                }}
              >
                {(Array.isArray(recents) ? recents : []).slice(0, uiRecentsDisplayCount).map((r, i) => {
                  const name = r?.name || "";
                  const id = r?.id;
                  const isFauxHighlighted = fauxFocus === "recents" && i === highlightRec;
                  return (
                    <div
                      key={id || `${name}_${i}`}
                      ref={(el) => { navRecRowRefs.current[i] = el; }}
                      onClick={() => { if (!isActivating && id) { setFauxFocus("recents"); setHighlightRec(i); onSelect({ id }); } }}
                      onMouseEnter={() => { if (suppressHoverRef.current) return; setFauxFocus("recents"); setHighlightRec(i); setHoverNavRecentId(id); }}
                      onMouseLeave={() => setHoverNavRecentId(null)}
                      style={{ ...rowStyle, background: isFauxHighlighted ? "rgba(0,120,212,0.12)" : (hoverNavRecentId === id && fauxFocus !== "recents" ? "rgba(0,120,212,0.07)" : "transparent") }}
                      role="button"
                    >
                      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <div title={name} style={{ flex: "1 1 auto", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{name}</div>
                      </div>
                    </div>
                  );
                })}
                {(Array.isArray(recents) ? recents : []).length === 0 && (
                  <div style={{ padding: "10px 12px", fontSize: 13, opacity: 0.75 }}>No recents yet.</div>
                )}
              </div>
            </div>
          </div>
        </>
      )}


      {activeTab === "Favorites" && (
        <>
          <div style={{ display: "flex", gap: 16, height: panelHeight, overflowX: "auto", overflowY: "hidden" }}>
            {/* Left: Search + Available (non-favorites) */}
            <div style={{ flex: "1 1 44%", minWidth: 240, paddingRight: 16, borderRight: "1px solid #d0d0d0", display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
              <div style={{ marginBottom: 10 }}>
                <input
                  autoFocus
                  ref={searchInputRef}
                  value={query}
                  onChange={(e) => setQuery(e.target.value)}
                  onBlur={() => requestSearchFocus("fav-search-blur")}
                  onKeyDown={(e) => {
                    try {
                      const key = e.key;
                      if (key === "Tab") {
                        e.preventDefault();
                        requestSearchFocus("tab");
                        return;
                      }
                      if (key === "ArrowDown") {
                        e.preventDefault();
                        // Mirror Navigation: move highlight through the available list (non-favorites)
                        const available = (Array.isArray(filtered) ? filtered : []).filter((x) => x && !isFavorite(x.id));
                        setHighlightIndex((prev) => Math.min((prev ?? -1) + 1, Math.max(available.length - 1, 0)));
                        return;
                      }
                      if (key === "ArrowUp") {
                        e.preventDefault();
                        setHighlightIndex((prev) => Math.max((prev ?? 0) - 1, 0));
                        return;
                      }
                      if (key === "Enter") {
                        e.preventDefault();
                        const available = (Array.isArray(filtered) ? filtered : []).filter((x) => x && !isFavorite(x.id));
                        const s = available[highlightIndex];
                        if (s?.id) addFavoriteLocal(s.id);
                        return;
                      }
                    } catch (e) {
                      // ignore
                    }
                  }}
                  placeholder="Search sheets…"
                  disabled={!!status && status !== "" && allSheets.length === 0}
                  style={{
                    width: "100%",
                    padding: "6px 8px",
                    fontSize: 12,
                    boxSizing: "border-box",
                    border: "1px solid rgba(0,0,0,0.2)",
                    borderRadius: 6,
                  }}
                />
              </div>

              <div
                style={{
                  border: "1px solid rgba(0,0,0,0.15)",
                  borderRadius: 6,
                  overflow: "hidden",
                  display: "flex",
                  flexDirection: "column",
                  flex: "1 1 auto",
                  minHeight: 0,
                  }}><div style={{ flex: "1 1 auto", minHeight: 0, overflowY: "auto", overscrollBehavior: "contain" }}>
                  {(Array.isArray(filtered) ? filtered : [])
                    .filter((s) => s && !isFavorite(s.id))
                    .map((s, i) => {
                      const isHovered = hoverFavTabAvailableId === s.id;
                      const isSel = favTabSelectedAvailableId === s.id;
                      const bg = (isSel || isHovered) ? "rgba(0,120,212,0.12)" : "transparent";
                      const boxShadow = isSel ? "inset 0 0 0 1px rgba(0,120,212,0.95)" : "none";
                      return (
                        <div
                          key={s.id}
                          onClick={() => {
                            if (isActivating) return;
                            setFavTabSelectedAvailableId(s.id);
                            setFavTabSelectedFavoriteId(null);
                            requestSearchFocus("fav-available-click");
                          }}
                          onDoubleClick={() => {
                            if (isActivating) return;
                            addFavoriteLocal(s.id);
                            requestSearchFocus("fav-available-dblclick");
                          }}
                          onMouseEnter={() => setHoverFavTabAvailableId(s.id)}
                          onMouseLeave={() => setHoverFavTabAvailableId(null)}
                          style={{
                            ...rowStyle,
                            background: bg,
                            boxShadow,
                          }}
                          role="button"
                          tabIndex={0}
                        >
                          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                            <div title={s.name} style={{ flex: "1 1 auto", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                              {s.name}
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  {(Array.isArray(filtered) ? filtered : []).filter((s) => s && !isFavorite(s.id)).length === 0 && (
                    <div style={{ padding: "10px 12px", fontSize: 13, opacity: 0.8 }}>
                      No matches.
                    </div>
                  )}
                </div>
              </div>
            </div>

            {/* Right: Favorites (top) + Controls (bottom, replaces Recents section) */}
            <div style={{ flex: "0 0 45%", minWidth: 220, display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
              {/* Favorites list */}
              <div style={{ marginBottom: 6 }}>
                <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 6, opacity: 0.85 }}>Favorites</div>
                <div
                  ref={favTabFavListRef}
                  style={{
                    height: favTabFavListHeight,
                    maxHeight: favTabFavListHeight,
                    minHeight: favTabFavListHeight,
                    overflowY: "auto",
                    overscrollBehavior: "contain",
                    border: "1px solid rgba(0,0,0,0.1)",
                    borderRadius: 6,
                  }}
                >
                  {(Array.isArray(favorites) ? favorites : []).map((f, i) => {
                    const name = f?.name || "";
                    const id = f?.id;
                    const isHovered = hoverFavTabFavoriteId === id;
                    const isSelected = favTabSelectedFavoriteId === id;
                    // Favorites tab favorites list: show a single highlight.
                    // - If a row is selected (clicked), highlight the selected row (needed for Up/Down).
                    // - If no selection, highlight follows mouse hover.
                    const bg = (isSelected || isHovered) ? "rgba(0,120,212,0.12)" : "transparent";
                    const boxShadow = isSelected ? "inset 0 0 0 1px rgba(0,120,212,0.95)" : "none";
                    return (
                      <div
                        key={id || `${name}_${i}`}
                        data-sheetid={id || ""}
                        onClick={() => {
                          if (isActivating) return;
                          if (id) setFavTabSelectedFavoriteId(id);
                          setFavTabSelectedAvailableId(null);
                          requestSearchFocus("fav-favorite-click");
}}
                        onDoubleClick={() => {
                          if (isActivating) return;
                          if (id) removeFavoriteLocal(id);
                          requestSearchFocus("fav-favorite-dblclick");
                        }}
                        onMouseEnter={() => id && setHoverFavTabFavoriteId(id)}
                        onMouseLeave={() => setHoverFavTabFavoriteId(null)}
                        style={{ ...rowStyle, background: bg, boxShadow }}
                        role="button"
                        tabIndex={0}
                      >
                        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                          <div style={{ width: 18, opacity: 0.75, textAlign: "right" }}>{i < 9 ? String(i + 1) : ""}</div>
                          <div title={name} style={{ flex: "1 1 auto", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{name}</div>
                        </div>
                      </div>
                    );
                  })}
                  {(Array.isArray(favorites) ? favorites : []).length === 0 && (
                    <div style={{ padding: "10px 12px", fontSize: 13, opacity: 0.75 }}>No favorites yet.</div>
                  )}
                </div>
              </div>

              {/* Controls block (mirrors where Recents was, but without Recents title) */}
              <div style={{ height: favTabBottomBlockHeight, maxHeight: favTabBottomBlockHeight, minHeight: favTabBottomBlockHeight, overflow: "visible", display: "flex", flexDirection: "column", justifyContent: "flex-end", paddingBottom: 8 }}>
                <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
                  <button
                    type="button"
                    disabled={!favTabSelectedFavoriteId || (Array.isArray(favorites) ? favorites : []).findIndex((x) => x?.id === favTabSelectedFavoriteId) <= 0}
                    onClick={() => moveFavoriteLocal(favTabSelectedFavoriteId, "up")}
                    style={{ flex: 1, padding: "6px 8px", fontSize: 12, borderRadius: 6, border: "1px solid rgba(0,0,0,0.2)", background: "white" }}
                  >
                    Up
                  </button>
                  <button
                    type="button"
                    disabled={
                      !favTabSelectedFavoriteId ||
                      (Array.isArray(favorites) ? favorites : []).findIndex((x) => x?.id === favTabSelectedFavoriteId) < 0 ||
                      (Array.isArray(favorites) ? favorites : []).findIndex((x) => x?.id === favTabSelectedFavoriteId) >= (Array.isArray(favorites) ? favorites : []).length - 1
                    }
                    onClick={() => moveFavoriteLocal(favTabSelectedFavoriteId, "down")}
                    style={{ flex: 1, padding: "6px 8px", fontSize: 12, borderRadius: 6, border: "1px solid rgba(0,0,0,0.2)", background: "white" }}
                  >
                    Down
                  </button>
                </div>

                <div style={{ textAlign: "center", fontSize: 14, fontWeight: 600, marginTop: 6, opacity: 0.85, userSelect: "none" }}>
                  ⇄&nbsp;&nbsp;&nbsp;Double-click to transfer&nbsp;&nbsp;&nbsp;⇄
                </div>
              </div>
            </div>
          </div>
        </>
      )}

      {activeTab === "Settings" && (
        <div style={{ height: panelHeight, overflowY: "auto", overflowX: "auto", paddingRight: 4, minWidth: 390 }}>
          <div style={{ minWidth: 520 }}>
          <div style={{ fontSize: 13, fontWeight: 800, margin: "2px 0 10px", opacity: 0.9 }}>Appearance</div>
          <div style={{ border: "1px solid rgba(0,0,0,0.12)", borderRadius: 10, padding: "10px 12px", marginBottom: 12 }}>
            <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 8, opacity: 0.9 }}>Row height</div>

            <div style={{ display: "flex", flexWrap: "wrap", gap: 14, alignItems: "center", fontSize: 12, opacity: 0.95 }}>
              {["Compact", "Standard", "Comfortable", "Expanded"].map((name) => (
                <label key={name} style={{ display: "flex", alignItems: "center", gap: 6, userSelect: "none" }}>
                  <input
                    type="radio"
                    name="rowHeightPreset_final"
                    checked={activePresetName === name}
                    onChange={() => {
                      const nextPreset = String(name);
                      rowHeightDirtyRef.current = true;
                      globalOptionsDirtyRef.current = true;
                      globalOptionsDirtyDesiredRef.current = {
                        oneDigitActivationEnabled: !!(globalOptions?.oneDigitActivationEnabled),
                        rowHeightPreset: nextPreset,
                      };
                      setGlobalOptions((prev) => ({ ...(prev || {}), rowHeightPreset: nextPreset }));
                    }}
                  />
                  {name}
                </label>
              ))}
            </div>
          </div>
          
          
          <div style={{ border: "1px solid rgba(0,0,0,0.12)", borderRadius: 10, padding: "10px 12px", marginBottom: 12 }}>
            <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 8, opacity: 0.9 }}>
              When space is limited, give more room to:
            </div>

            <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 8 }}>
              <div style={{ width: 66, fontSize: 12, opacity: 0.85 }}>Favorites</div>
              <input
                type="range"
                min={20}
                max={80}
                step={5}
                value={100 - favPercentEffective}
                onChange={(e) => {
                  const v = Math.min(80, Math.max(20, Number(e.target.value) || 20));
                  const nextFav = Math.min(80, Math.max(20, Math.round(100 - v)));
                  uiSettingsDirtyRef.current = true;
                  uiSettingsDirtyDesiredRef.current = {
                    favPercentManual: nextFav,
                    recentsDisplayCount: Math.min(MAX_RECENTS, Math.max(1, Math.round(uiRecentsDisplayCount))),
                  };
                  setUiFavPercentManual(nextFav);
                }}
                style={{ flex: "1 1 auto" }}
              />
              <div style={{ width: 66, fontSize: 12, opacity: 0.85, textAlign: "right" }}>Recents</div>
              <div style={{ width: 170, fontSize: 12, opacity: 0.85, textAlign: "right" }}>
                Favorites {favPercentEffective}% / Recents {recPercentEffective}%
              </div>
            </div>
          </div>

          <div style={{ fontSize: 13, fontWeight: 800, margin: "12px 0 10px", opacity: 0.9 }}>Behavior</div>
          <div style={{ border: "1px solid rgba(0,0,0,0.12)", borderRadius: 10, padding: "10px 12px", marginBottom: 12 }}>
            <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 8, opacity: 0.9 }}>List ordering</div>

            <div style={{ display: "flex", flexDirection: "column", gap: 10, fontSize: 12, opacity: 0.95 }}>
              <div>
                <label style={{ display: "flex", gap: 8, alignItems: "center", cursor: "pointer" }}>
                  <input
                    type="radio"
                    name="baselineOrder"
                    checked={String(globalOptions?.baselineOrder || "workbook") !== "alpha"}
                    onChange={() => {
                      setGlobalOptions((prev) => ({ ...(prev || {}), baselineOrder: "workbook" }));
                      schedulePersistUiSettings("baselineOrder");
                    }}
                  />
                  Workbook order
                </label>

                  <label style={{ display: "flex", gap: 8, alignItems: "center", cursor: "pointer" }}>
                    <input
                      type="radio"
                      name="baselineOrder"
                      checked={String(globalOptions?.baselineOrder || "workbook") === "alpha"}
                      onChange={() => {
                        setGlobalOptions((prev) => ({ ...(prev || {}), baselineOrder: "alpha" }));
                        schedulePersistUiSettings("baselineOrder");
                      }}
                    />
                    Alphabetical
                  </label>
              </div>
            </div>
          </div>

          <div style={{ border: "1px solid rgba(0,0,0,0.12)", borderRadius: 10, padding: "10px 12px", marginBottom: 12 }}>
            <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 8, opacity: 0.9 }}>Frequent on top</div>

            <label style={{ display: "flex", alignItems: "flex-start", gap: 8, fontSize: 12, opacity: (PREMIUM_FREQ_BUMP || devPremiumRef.current) ? 0.95 : 0.45, userSelect: "none", cursor: (PREMIUM_FREQ_BUMP || devPremiumRef.current) ? "pointer" : "default" }}>
              <input
                type="checkbox"
                checked={!!(globalOptions?.frequentOnTop)}
                disabled={!(PREMIUM_FREQ_BUMP || devPremiumRef.current)}
                onChange={(e) => {
                  if (!(PREMIUM_FREQ_BUMP || devPremiumRef.current)) return;
                  const nextEnabled = !!e.target.checked;
                  setGlobalOptions((prev) => ({ ...(prev || {}), frequentOnTop: nextEnabled }));
                  schedulePersistUiSettings("frequentOnTop");
                }}
                style={{ marginTop: 2 }}
              />
              <div>
                <div style={{ fontWeight: 600 }}>Promote frequently used sheets</div>
                <div style={{ marginTop: 4, opacity: 0.85 }}>When searching, heavily-used sheets appear at top of search results.</div>
                {!(PREMIUM_FREQ_BUMP || devPremiumRef.current) && (
                  <div style={{ marginTop: 4, color: "rgba(0,0,0,0.45)", fontStyle: "italic" }}>Premium feature.</div>
                )}
              </div>
            </label>
          </div>

          <div style={{ border: "1px solid rgba(0,0,0,0.12)", borderRadius: 10, padding: "10px 12px", marginBottom: 12 }}>
            <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 8, opacity: 0.9 }}>Recents</div>

            <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, opacity: 0.9 }}>
              <span>Show</span>
              <input
                type="number"
                min={1}
                max={MAX_RECENTS}
                value={uiRecentsDisplayCount}
                onChange={(e) => {
                  const v = Math.min(MAX_RECENTS, Math.max(1, Number(e.target.value) || 1));
                  const nextCnt = Math.min(MAX_RECENTS, Math.max(1, Math.round(v)));
                  uiSettingsDirtyRef.current = true;
                  uiSettingsDirtyDesiredRef.current = {
                    favPercentManual: Math.min(80, Math.max(20, Math.round(uiFavPercentManual))),
                    recentsDisplayCount: nextCnt,
                  };
                  setUiRecentsDisplayCount(nextCnt);
                }}
                style={{ width: 64, padding: "2px 6px", fontSize: 12, border: "1px solid rgba(0,0,0,0.25)", borderRadius: 6 }}
              />
              <span>items</span>
            </div>
          </div>

<div style={{ border: "1px solid rgba(0,0,0,0.12)", borderRadius: 10, padding: "10px 12px", marginBottom: 12 }}>
            <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 8, opacity: 0.9 }}>One-digit activation</div>

            <label style={{ display: "flex", alignItems: "flex-start", gap: 8, fontSize: 12, opacity: isReadOnly ? 0.45 : 0.95, userSelect: "none", cursor: isReadOnly ? "default" : "pointer" }}>
              <input
                type="checkbox"
                checked={!!(globalOptions?.oneDigitActivationEnabled)}
                disabled={isReadOnly}
                onChange={(e) => {
                  if (isReadOnly) return;
                  const nextEnabled = !!e.target.checked;
                  globalOptionsDirtyRef.current = true;
                  // Capture desired globals so we can ignore stale stateData until parent echoes the same values back.
                  globalOptionsDirtyDesiredRef.current = {
                    oneDigitActivationEnabled: nextEnabled,
                    rowHeightPreset: String(globalOptions?.rowHeightPreset || "Standard"),
                  };
                  setGlobalOptions((prev) => ({ ...(prev || {}), oneDigitActivationEnabled: nextEnabled }));
                }}
                style={{ marginTop: 2 }}
              />
              <div>
                <div style={{ fontWeight: 600 }}>Enable one-digit activation for this workbook</div>
                <div style={{ marginTop: 4, opacity: 0.85 }}>Jump instantly to a Favorite by typing a single digit (1–9, 0).</div>
                <div style={{ marginTop: 4, opacity: 0.85 }}>Tip: To search for numbers (e.g. 2024), start the search with a space.</div>
                {isReadOnly && (
                  <div style={{ marginTop: 4, color: "rgba(0,0,0,0.45)", fontStyle: "italic" }}>Not available in read-only workbooks.</div>
                )}
              </div>
            </label>
          </div>

<div style={{ border: "1px solid rgba(0,0,0,0.12)", borderRadius: 10, padding: "10px 12px", marginBottom: 12 }}>
            <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 8, opacity: 0.9 }}>Quick Return</div>

            <label style={{ display: "flex", alignItems: "flex-start", gap: 8, fontSize: 12, opacity: 0.95, userSelect: "none", cursor: "pointer" }}>
              <input
                type="checkbox"
                checked={enableQuickReturn}
                onChange={(e) => {
                  const nextEnabled = !!e.target.checked;
                  setEnableQuickReturn(nextEnabled);
                  try {
                    if (Office?.context?.ui?.messageParent) {
                      Office.context.ui.messageParent(JSON.stringify({ type: "setEnableQuickReturn", enabled: nextEnabled }));
                    }
                  } catch (err) {
                    console.error("messageParent(setEnableQuickReturn) failed:", err);
                  }
                }}
                style={{ marginTop: 2 }}
              />
              <div>
                <div style={{ fontWeight: 600 }}>Enable quick return</div>
                <div style={{ marginTop: 4, opacity: 0.85 }}>To return to your previous sheet, open JumpTo and simply press Enter — available when you&rsquo;re still on the sheet you jumped to.</div>
              </div>
            </label>
          </div>

          
          </div>
        </div>
      )}

      </div>
      {/* Global actions (outside tabs) */}
      <div ref={footerRef} style={{ display: "flex", justifyContent: "flex-end", alignItems: "center", marginTop: 8, paddingTop: 8, borderTop: "1px solid #e0e0e0" }}>
        <button
          type="button"
          onClick={() => {
            try {
              if (window.Office?.context?.ui?.messageParent) {
                onCancel();
              } else {
                window.close?.();
              }
            } catch (e) {
              console.error("Close failed:", e);
              window.close?.();
            }
          }}
          style={{
            padding: "6px 14px",
            fontSize: 12,
            border: "1px solid #c8c8c8",
            borderRadius: 6,
            background: "#f5f5f5",
            cursor: "pointer",
          }}
        >
          Close
        </button>
      </div>
    </div>
  );
}

const rootEl = document.getElementById("root");
function boot() {
  if (!rootEl) return;
  createRoot(rootEl).render(<DialogApp />);
}

// Office.js can load after our bundle in some dialog webviews (script loading race).
// To avoid calling Office APIs too early, wait briefly for Office to appear, then gate on Office.onReady.
function waitForOfficeGlobal(timeoutMs = 4000, pollMs = 25) {
  return new Promise((resolve) => {
    const start = Date.now();
    const tick = () => {
      if (window.Office) return resolve(true);
      if (Date.now() - start >= timeoutMs) return resolve(false);
      window.setTimeout(tick, pollMs);
    };
    tick();
  });
}

(async () => {
  try {
    const hasOffice = await waitForOfficeGlobal();
    if (hasOffice && typeof Office.onReady === "function") {
      Office.onReady(() => boot());
    } else {
      // Dev-friendly: still render so the dialog page can be opened in a normal browser.
      boot();
    }
  } catch (e) {
    // As a last resort, render a UI so we can surface the error.
    boot();
  }
})();

// DEBUG: receive persistence diagnostics from parent (temporary)
try {
  if (typeof Office !== "undefined" && Office.context && Office.context.ui && Office.context.ui.addHandlerAsync) {
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, function (arg) {
      try {
        const msg = JSON.parse(arg.message);
      } catch (e) {
        // ignore non-JSON
      }
    });
  }
} catch (e) {
  // no-op
}