import React, { useEffect, useMemo, useRef, useState } from "react";
import { MAX_RECENTS } from "../shared/constants";

import { createRoot } from "react-dom/client";

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
  } catch {
    return null;
  }
}

function TabButton({ label, active, onClick }) {
  return (
    <button
      type="button"
      onClick={onClick}
      style={{
        appearance: "none",
        background: "transparent",
        border: "none",
        padding: "8px 12px",
        margin: 0,
        cursor: "pointer",
        fontFamily: "Segoe UI, Arial, sans-serif",
        fontSize: 13,
        fontWeight: active ? 600 : 400,
        color: "#111",
        borderBottom: active ? "2px solid #0078d4" : "2px solid transparent",
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


function DialogApp() {
  const [allSheets, setAllSheets] = useState([]);
  const [favorites, setFavorites] = useState([]);
  const [recents, setRecents] = useState([]);
  const [globalOptions, setGlobalOptions] = useState({ oneDigitActivationEnabled: true, rowHeightPreset: "Standard", baselineOrder: "workbook", frequentOnTop: true });
  const [query, setQuery] = useState("");
  const [status, setStatus] = useState("Loading…");
  const [isActivating, setIsActivating] = useState(false);
  const [initError, setInitError] = useState("");
  const [activeTab, setActiveTab] = useState("Navigation");
  
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
  const [uiRecentsDisplayCount, setUiRecentsDisplayCount] = useState(10); // 1..MAX_RECENTS
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

  const [highlightIndex, setHighlightIndex] = useState(0);
  const requestedRef = useRef(false);
  const timeoutIdRef = useRef(null);
  const statusRef = useRef("Loading…");
  const sheetsLenRef = useRef(0);
  const searchInputRef = useRef(null);
  const listRowRefs = useRef([]);
  const focusTimersRef = useRef([]);
  const parentReadyRef = useRef(false);
  const uiSettingsReadyRef = useRef(false);
  useEffect(() => { favoritesRef.current = favorites; }, [favorites]);

  useEffect(() => { statusRef.current = status; }, [status]);
  useEffect(() => { sheetsLenRef.current = allSheets.length; }, [allSheets]);

  const requestSearchFocus = (reason = "") => {
    // Office dialog webviews can be finicky with focus timing. Be defensive and never throw.
    // Cancel any existing scheduled focus attempts.
    try {
      (focusTimersRef.current || []).forEach((t) => window.clearTimeout(t));
    } catch {
      // ignore
    }
    focusTimersRef.current = [];

    const tryFocus = () => {
      const el = searchInputRef.current;
      if (!el || typeof el.focus !== "function") return;
      try {
        el.focus();
      } catch {
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
      } catch {
        // ignore
      }
    };
    const onUnhandled = (evt) => {
      try {
        const reason = evt?.reason;
        const msg = reason?.message || String(reason || "Unhandled promise rejection");
        console.error("[JumpToSheet][Dialog] unhandledrejection:", msg, evt);
        setInitError((prev) => prev || msg);
      } catch {
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
      } catch {
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
      } catch {
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
    } catch {}
    return () => {
      window.removeEventListener("resize", onResize);
      try {
        if (ro) ro.disconnect();
      } catch {}
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
      } catch {
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
        console.log(`[JT][build 37] dialog ready`, window.location.href);
      } catch { /* ignore */ }
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
            const state = msg.state || {};
            const sheets = Array.isArray(state.sheets) ? state.sheets : [];
            setAllSheets(sheets);
            setFavorites(Array.isArray(state.favorites) ? state.favorites : []);
            setRecents(Array.isArray(state.recents) ? state.recents : []);
            setGlobalOptions(state.global || { oneDigitActivationEnabled: true, rowHeightPreset: "Standard", baselineOrder: "workbook", frequentOnTop: true });
            // UI settings (persisted per-user)
            try {
              const ui = state.settings || {};              const favPct = Number.isFinite(Number(ui.favPercentManual)) ? Number(ui.favPercentManual) : 50;
              const recCnt = Number.isFinite(Number(ui.recentsDisplayCount)) ? Number(ui.recentsDisplayCount) : 10;
              setUiFavPercentManual(Math.min(80, Math.max(20, Math.round(favPct))));
              setUiRecentsDisplayCount(Math.min(MAX_RECENTS, Math.max(1, Math.round(recCnt))));
            } catch (e) {
              // ignore
            }
            uiSettingsReadyRef.current = true;
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
        try { requestSheets(); } catch (e) {}
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
    if (q) items = items.filter((s) => (s?.name || "").toLowerCase().includes(q));

    // Baseline order
    const baseline = (globalOptions?.baselineOrder || "workbook");
    if (baseline === "alpha") {
      items.sort((a, b) => (a?.name || "").localeCompare(b?.name || ""));
    } else {
      items.sort((a, b) => Number(a?.orderIndex || 0) - Number(b?.orderIndex || 0));
    }

    // Frequent-on-top: move only highest tier present to the top
    if (globalOptions?.frequentOnTop) {
      const tiers = items.map((s) => computeTier(s?.freq || 0));
      const maxTier = tiers.length ? Math.max(...tiers) : 0;
      if (maxTier > 0) {
        const high = [];
        const rest = [];
        for (let i = 0; i < items.length; i++) {
          if (tiers[i] === maxTier) high.push(items[i]);
          else rest.push(items[i]);
        }
        items = [...high, ...rest];
      }
    }

    return items;
  }, [allSheets, query, globalOptions]);

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
    if (!sheetId) return;
    setFavorites((prev) => {
      const arr = Array.isArray(prev) ? prev : [];
      if (arr.some((x) => x?.id === sheetId)) return arr;
      const s = (Array.isArray(allSheets) ? allSheets : []).find((x) => x?.id === sheetId);
      const name = s?.name || "";
      return [...arr, { id: sheetId, name }];
    });
    setFavTabSelectedFavoriteId(sheetId);
    setFavTabSelectedAvailableId(null);
    favTabPendingScrollIdRef.current = sheetId;
    schedulePersistFavorites("add");
  };

  const removeFavoriteLocal = (sheetId) => {
    if (!sheetId) return;
    setFavorites((prev) => (Array.isArray(prev) ? prev : []).filter((x) => x?.id !== sheetId));
    if (favTabSelectedFavoriteId === sheetId) setFavTabSelectedFavoriteId(null);
    schedulePersistFavorites("remove");
  };

  const moveFavoriteLocal = (sheetId, direction) => {
    if (!sheetId) return;
    if (direction !== "up" && direction !== "down") return;
    setFavorites((prev) => {
      const arr = Array.isArray(prev) ? prev.slice() : [];
      const idx = arr.findIndex((x) => x?.id === sheetId);
      if (idx < 0) return arr;
      const to = direction === "up" ? idx - 1 : idx + 1;
      if (to < 0 || to >= arr.length) return arr;
      const [item] = arr.splice(idx, 1);
      arr.splice(to, 0, item);
      return arr;
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
    if (!uiSettingsReadyRef.current) return;
    if (uiSettingsPersistTimerRef.current) {
      clearTimeout(uiSettingsPersistTimerRef.current);
    }
    uiSettingsPersistTimerRef.current = setTimeout(() => {
      uiSettingsPersistTimerRef.current = null;
      try {
        sendSetUiSettingsToParent({          favPercentManual: Math.min(80, Math.max(20, Math.round(uiFavPercentManual))),
          recentsDisplayCount: Math.min(MAX_RECENTS, Math.max(1, Math.round(uiRecentsDisplayCount))),        });
      } catch {
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
      sendSetUiSettingsToParent({        favPercentManual: Math.min(80, Math.max(20, Math.round(uiFavPercentManual))),
        recentsDisplayCount: Math.min(MAX_RECENTS, Math.max(1, Math.round(uiRecentsDisplayCount))),        });
    } catch {
      // ignore
    }
  };

  const schedulePersistGlobalOptions = (reason) => {
    if (globalOptionsPersistTimerRef.current) {
      clearTimeout(globalOptionsPersistTimerRef.current);
    }
    globalOptionsPersistTimerRef.current = setTimeout(() => {
      globalOptionsPersistTimerRef.current = null;
      try {
        const preset = String(globalOptions?.rowHeightPreset || "Standard");

        try {

          if (Office?.context?.ui?.messageParent) {

            Office.context.ui.messageParent(JSON.stringify({ type: "setRowHeightPreset", preset }));

          }

        } catch (err) {

          console.error("messageParent(setRowHeightPreset) failed:", err);

        }} catch {
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

          Office.context.ui.messageParent(JSON.stringify({ type: "setRowHeightPreset", preset }));

        }

      } catch (err) {

        console.error("messageParent(setRowHeightPreset) failed:", err);

      }} catch {
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
    } catch {
      // ignore
    }
  }, []);

  // Persist UI settings when they change (debounced).
  useEffect(() => {
    if (!parentReadyRef.current) return;
    schedulePersistUiSettings("ui-change");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [uiFavPercentManual, uiRecentsDisplayCount]);

  // Persist global options when they change (debounced).
  useEffect(() => {
    if (!parentReadyRef.current) return;
    schedulePersistGlobalOptions("global-change");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [globalOptions?.rowHeightPreset]);

  // Expose flush for Save & Close
  useEffect(() => {
    window.flushPersistUiSettingsNow = flushPersistUiSettingsNow;
    return () => { try { delete window.flushPersistUiSettingsNow; } catch {} };
  }, [uiFavPercentManual, uiRecentsDisplayCount]);

  useEffect(() => {
    window.flushPersistGlobalOptionsNow = flushPersistGlobalOptionsNow;
    return () => { try { delete window.flushPersistGlobalOptionsNow; } catch {} };
  }, [globalOptions?.rowHeightPreset]);

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
        } catch {
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
    } catch {
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
        const ids = (Array.isArray(favoritesRef.current) ? favoritesRef.current : []).map((x) => x?.id).filter(Boolean);
        sendSetFavoritesToParent(ids);
        favDirtyRef.current = false;
      } catch {
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
      const ids = (Array.isArray(favoritesRef.current) ? favoritesRef.current : []).map((x) => x?.id).filter(Boolean);
      sendSetFavoritesToParent(ids);
    } catch {
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
    };

    const favoritesIds = (Array.isArray(favoritesRef.current) ? favoritesRef.current : [])
      .map((x) => x?.id)
      .filter(Boolean);

    const rowHeightPreset = String(globalOptions?.rowHeightPreset || "Standard");

    return { uiSettings, favorites: favoritesIds, rowHeightPreset };
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
    Office.context.ui.messageParent(JSON.stringify({ type: "cancel", snapshot }));
  } catch {
    // ignore
  }
};

  // Listbox-like navigation: default highlight is first row after load/filter.
useEffect(() => {
  if (activeTab !== "Navigation") return;
  setHighlightIndex(0);
  // Do NOT re-select the search text on every keystroke; that would cause each new character to replace the previous.
  requestSearchFocus("resetHighlight");
  // eslint-disable-next-line react-hooks/exhaustive-deps
}, [filtered.length, activeTab]);

// Keep the highlighted row visible when navigating with arrow keys.
useEffect(() => {
  if (activeTab !== "Navigation") return;
  const el = listRowRefs.current?.[highlightIndex];
  if (el && typeof el.scrollIntoView === "function") {
    try {
      el.scrollIntoView({ block: "nearest" });
    } catch {
      // ignore
    }
  }
}, [highlightIndex, activeTab]);

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

      <div ref={tabsRef}
        style={{
          display: "flex",
          borderBottom: "1px solid rgba(0,0,0,0.15)",
          marginBottom: 10,
          marginTop: 2,
        }}
        role="tablist"
        aria-label="JumpTo tabs"
      >
        <TabButton label="Navigation" active={activeTab === "Navigation"} onClick={() => setActiveTab("Navigation")} />
        <TabButton label="Favorites" active={activeTab === "Favorites"} onClick={() => setActiveTab("Favorites")} />
        <TabButton label="Settings" active={activeTab === "Settings"} onClick={() => setActiveTab("Settings")} />
      </div>

      <div ref={bodyRef} style={{ flex: "1 1 auto", overflow: "hidden" }}>

      {activeTab === "Navigation" && (
        <>
          <div style={{ display: "flex", gap: 16, height: panelHeight, overflow: "hidden" }}>
            {/* Left: Search + All results */}
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
                      // Keep focus in search box; allow navigation + activation like a VBA listbox.
                      if (key === "Tab") {
                        e.preventDefault();
                        requestSearchFocus("tab");
                        return;
                      }
                      if (key === "ArrowDown") {
                        e.preventDefault();
                        setHighlightIndex((prev) => {
                          const max = Math.max(0, (filtered?.length || 0) - 1);
                          return Math.min(max, prev + 1);
                        });
                        return;
                      }
                      if (key === "ArrowUp") {
                        e.preventDefault();
                        setHighlightIndex((prev) => Math.max(0, prev - 1));
                        return;
                      }
                      if (key === "Enter") {
                        e.preventDefault();
                        const idx = Math.max(0, Math.min((filtered?.length || 1) - 1, highlightIndex));
                        const s = filtered?.[idx];
                        if (s) onSelect(s);
                        return;
                      }
                      const mods = e.altKey || e.ctrlKey || e.metaKey;
                      const oneDigit = globalOptions?.oneDigitActivationEnabled;
                      const q = query || "";
                      const leadingSpace = q.startsWith(" ");

                      // One-digit activation: only when search box is empty, no modifiers, and no leading space.
                      if (oneDigit && !mods && !leadingSpace && q === "" && key >= "0" && key <= "9") {
                        const idx = key === "0" ? 9 : (Number(key) - 1);
                        const fav = favorites?.[idx];
                        if (fav?.id) {
                          e.preventDefault();
                          onSelect(fav);
                          return;
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
                    } catch {
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
                    border: "1px solid rgba(0,0,0,0.1)",
                    borderRadius: 6,
                  }}
                >
                  {filtered.map((s, i) => (
                    <div
                      key={s.id || s.name}
                      ref={(el) => { listRowRefs.current[i] = el; }}
                      onMouseEnter={() => { try { setHighlightIndex(i); } catch {} }}
                      onClick={() => { if (!isActivating) { try { setHighlightIndex(i); } catch {} onSelect(s); } }}
                      style={{
                        ...rowStyle,
                        background: i === highlightIndex ? "rgba(0,120,212,0.12)" : "transparent",
                      }}

                      role="button"
                      tabIndex={0}
                      onKeyDown={(e) => {
                        if (isActivating) return;
                        if (e.key === "Enter" || e.key === " ") onSelect(s);
                      }}
                    >
                      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <div style={{ flex: "1 1 auto", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{s.name}</div>
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
                style={{
                  flex: "0 1 auto",
                  height: navTabFavListHeight,
                  maxHeight: navTabFavListHeight,
                  minHeight: 0,
                  overscrollBehavior: "contain",
                  overflowY: "auto",
                  overflowX: "hidden",
                  boxSizing: "border-box",
                  border: "1px solid rgba(0,0,0,0.1)",
                  borderRadius: 6,
                  marginBottom: 6,
                }}>
                {(Array.isArray(favorites) ? favorites : []).map((f, i) => {
                  const slot = i < 9 ? String(i + 1) : i === 9 ? "0" : "-";
                  const name = f?.name || "";
                  const id = f?.id;
                  return (
                    <div
                      key={id || `${name}_${i}`}
                      onClick={() => !isActivating && id && onSelect({ id })}
                      onMouseEnter={() => setHoverNavFavoriteId(id)}
                      onMouseLeave={() => setHoverNavFavoriteId(null)}
                      style={{ ...rowStyle, background: (hoverNavFavoriteId === id ? "rgba(0,120,212,0.10)" : "transparent") }}
                      role="button"
                      tabIndex={0}
                      onKeyDown={(e) => {
                        if (isActivating) return;
                        if (e.key === "Enter" || e.key === " ") id && onSelect({ id });
                      }}
                    >
                      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <div style={{ width: 18, opacity: 0.75, textAlign: "right" }}>{slot}</div>
                        <div style={{ flex: "1 1 auto", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{name}</div>
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
                style={{
                  flex: "0 1 auto",
                  height: navTabRecListHeight,
                  maxHeight: navTabRecListHeight,
                  minHeight: 0,
                  overscrollBehavior: "contain",
                  overflowY: "auto",
                  overflowX: "hidden",
                  boxSizing: "border-box",
                  border: "1px solid rgba(0,0,0,0.1)",
                  borderRadius: 6,
                }}
              >
                {(Array.isArray(recents) ? recents : []).slice(0, uiRecentsDisplayCount).map((r, i) => {
                  const name = r?.name || "";
                  const id = r?.id;
                  const fav = isFavorite(id);
                  return (
                    <div
                      key={id || `${name}_${i}`}
                      onClick={() => !isActivating && id && onSelect({ id })}
                      onMouseEnter={() => setHoverNavRecentId(id)}
                      onMouseLeave={() => setHoverNavRecentId(null)}
                      style={{ ...rowStyle, background: (hoverNavRecentId === id ? "rgba(0,120,212,0.10)" : "transparent") }}
                      role="button"
                    >
                      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <div style={{ flex: "1 1 auto", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{name}</div>
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
          <div style={{ display: "flex", gap: 16, height: panelHeight, overflow: "hidden" }}>
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
                    } catch {
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
                            <div style={{ flex: "1 1 auto", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
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
                          <div style={{ flex: "1 1 auto", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{name}</div>
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
        <div style={{ height: panelHeight, overflow: "hidden", paddingRight: 4 }}>
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
                  setUiFavPercentManual(100 - v);
                }}
                style={{ flex: "1 1 auto" }}
              />
              <div style={{ width: 66, fontSize: 12, opacity: 0.85, textAlign: "right" }}>Recents</div>
              <div style={{ width: 170, fontSize: 12, opacity: 0.85, textAlign: "right" }}>
                Favorites {favPercentEffective}% / Recents {recPercentEffective}%
              </div>
            </div>
          </div>
          
          <div style={{ border: "1px solid rgba(0,0,0,0.12)", borderRadius: 10, padding: "10px 12px", marginBottom: 12 }}>
            <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 8, opacity: 0.9 }}>Row height</div>

            <div style={{ display: "flex", flexWrap: "wrap", gap: 14, alignItems: "center", fontSize: 12, opacity: 0.95 }}>
              {["Compact", "Standard", "Comfortable", "Expanded"].map((name) => (
                <label key={name} style={{ display: "flex", alignItems: "center", gap: 6, userSelect: "none" }}>
                  <input
                    type="radio"
                    name="rowHeightPreset_final"
                    checked={activePresetName === name}
                    onChange={() => setGlobalOptions((prev) => ({ ...(prev || {}), rowHeightPreset: name }))}
                  />
                  {name}
                </label>
              ))}
            </div>
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
                  setUiRecentsDisplayCount(v);
                }}
                style={{ width: 64, padding: "2px 6px", fontSize: 12, border: "1px solid rgba(0,0,0,0.25)", borderRadius: 6 }}
              />
              <span>items</span>
            </div>
          </div>
        </div>
      )}

      </div>
      {/* Global actions (outside tabs) */}
      <div ref={footerRef} style={{ display: "flex", justifyContent: "flex-end", marginTop: 8, paddingTop: 8, borderTop: "1px solid #e0e0e0" }}>
        <button
          type="button"
          onClick={() => {
            try {
              if (window.Office?.context?.ui?.messageParent) {
                window.flushPersistFavoritesNow?.("close");
                window.flushPersistUiSettingsNow?.("close");
                window.flushPersistGlobalOptionsNow?.("close");
                Office.context.ui.messageParent(JSON.stringify({ type: "cancel", uiSettings: {      favPercentManual: Math.min(80, Math.max(20, Math.round(uiFavPercentManual))),
      recentsDisplayCount: Math.min(MAX_RECENTS, Math.max(1, Math.round(uiRecentsDisplayCount))),    }}));
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