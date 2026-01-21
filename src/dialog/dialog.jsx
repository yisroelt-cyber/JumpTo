
// JumpTo: harden dialog startup diagnostics
if (typeof window !== "undefined") {
  window.addEventListener("error", (e) => {
    // eslint-disable-next-line no-console
    console.error("JumpTo dialog startup error:", e?.error || e?.message || e);
    try {
      const el = document.getElementById("jumpto-startup-error") || (() => {
        const d = document.createElement("pre");
        d.id = "jumpto-startup-error";
        d.style.cssText = "white-space:pre-wrap;background:#fff3cd;border:1px solid #ffeeba;padding:10px;margin:10px;font-size:12px;color:#856404;";
        document.body.prepend(d);
        return d;
      })();
      el.textContent = String(e?.error?.stack || e?.error || e?.message || e);
    } catch {}
  });
  window.addEventListener("unhandledrejection", (e) => {
    // eslint-disable-next-line no-console
    console.error("JumpTo dialog unhandled rejection:", e?.reason || e);
  });
  // eslint-disable-next-line no-console
  console.log("JumpTo: dialog.jsx loaded");
}

import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";

/* global Office */

const JT_BUILD = "37";


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

function DialogApp() {
  const [allSheets, setAllSheets] = useState([]);
  const [favorites, setFavorites] = useState([]);
  const [recents, setRecents] = useState([]);
  const [globalOptions, setGlobalOptions] = useState({ oneDigitActivationEnabled: true, rowHeightPreset: "Compact", baselineOrder: "workbook", frequentOnTop: true });
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
const [highlightIndex, setHighlightIndex] = useState(0);
  const requestedRef = useRef(false);
  const timeoutIdRef = useRef(null);
  const statusRef = useRef("Loading…");
  const sheetsLenRef = useRef(0);
  const searchInputRef = useRef(null);
  const listRowRefs = useRef([]);
  const focusTimersRef = useRef([]);

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
    const parentReadyRef = { current: false };

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
            if (activeTab === "Navigation") requestSearchFocus("parentReady");
            return;
          }

          if (msg.type === "stateData") {
            const state = msg.state || {};
            const sheets = Array.isArray(state.sheets) ? state.sheets : [];
            setAllSheets(sheets);
            setFavorites(Array.isArray(state.favorites) ? state.favorites : []);
            setRecents(Array.isArray(state.recents) ? state.recents : []);
            setGlobalOptions(state.global || { oneDigitActivationEnabled: true, rowHeightPreset: "Compact", baselineOrder: "workbook", frequentOnTop: true });
            setStatus(sheets.length ? "" : "No visible worksheets found.");

            // Re-assert focus after data arrives (this is the moment users start typing).
            if (activeTab === "Navigation") requestSearchFocus("sheetsData");
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
  };

  const removeFavoriteLocal = (sheetId) => {
    if (!sheetId) return;
    setFavorites((prev) => (Array.isArray(prev) ? prev : []).filter((x) => x?.id !== sheetId));
    if (favTabSelectedFavoriteId === sheetId) setFavTabSelectedFavoriteId(null);
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
  };
const rowStyle = {
    padding: "2px 10px",
    fontSize: 12,
    lineHeight: "16px",
    cursor: isActivating ? "default" : "pointer",
    borderBottom: "1px solid rgba(0,0,0,0.06)",
    userSelect: "none",
    opacity: isActivating ? 0.65 : 1,
  };



  
const onSelect = (sheet) => {
  if (!sheet || isActivating) return;
  const sheetId = typeof sheet === "string" ? sheet : sheet.id;
  if (!sheetId) return;
  setIsActivating(true);
  setStatus("Loading sheet…");
  try {
    Office.context.ui.messageParent(JSON.stringify({ type: "selectSheet", sheetId }));
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
    Office.context.ui.messageParent(JSON.stringify({ type: "cancel" }));
  } catch {
    // ignore
  }
};

return (
    <div style={{ fontFamily: "Segoe UI, Arial, sans-serif", padding: 14 }}>
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

      <div
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

      {activeTab === "Navigation" && (
        <>
          <div style={{ display: "flex", gap: 16 }}>
            {/* Left: Search + All results */}
            <div style={{ flex: "1 1 44%", minWidth: 240, paddingRight: 16, borderRight: "1px solid #d0d0d0" }}>
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
                    maxHeight: 300,
              minHeight: 300,
                    overflowY: "auto",
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
                        padding: "2px 10px",
                        fontSize: 12,
                        lineHeight: "16px",
                        cursor: isActivating ? "default" : "pointer",
                        borderBottom: "1px solid rgba(0,0,0,0.06)",
                        userSelect: "none",
                        opacity: isActivating ? 0.65 : 1,
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
            <div style={{ flex: "0 0 45%", minWidth: 220 }}>
              <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 6, opacity: 0.85 }}>Favorites</div>
              <div
                style={{
                  maxHeight: 150,
                  minHeight: 150,
                  overflowY: "auto",
                  overscrollBehavior: "contain",
                  border: "1px solid rgba(0,0,0,0.1)",
                  borderRadius: 6,
                  marginBottom: 12,
                }}
              >
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

              <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 6, opacity: 0.85 }}>Recents</div>
              <div
                style={{
                  maxHeight: 150,
                  overflowY: "auto",
                  overscrollBehavior: "contain",
                  border: "1px solid rgba(0,0,0,0.1)",
                  borderRadius: 6,
                }}
              >
                {(Array.isArray(recents) ? recents : []).map((r, i) => {
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
          <div style={{ display: "flex", gap: 16 }}>
            {/* Left: Search + Available (non-favorites) */}
            <div style={{ flex: "1 1 44%", minWidth: 240, paddingRight: 16, borderRight: "1px solid #d0d0d0" }}>
              <div style={{ marginBottom: 10 }}>
                <input
                  autoFocus
                  ref={searchInputRef}
                  value={query}
                  onChange={(e) => setQuery(e.target.value)}
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
                }}
              >
                <div style={{ maxHeight: 300, minHeight: 300, overflowY: "auto", overscrollBehavior: "contain" }}>
                  {(Array.isArray(filtered) ? filtered : [])
                    .filter((s) => s && !isFavorite(s.id))
                    .map((s, i) => {
                      const isHovered = hoverFavTabAvailableId === s.id;
                      const isSel = favTabSelectedAvailableId === s.id;
                      const bg = isSel ? "rgba(0,120,212,0.12)" : isHovered ? "rgba(0,120,212,0.10)" : "transparent";
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
            <div style={{ flex: "0 0 45%", minWidth: 220, display: "flex", flexDirection: "column", gap: 14 }}>
              {/* Favorites list */}
              <div>
                <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 6, opacity: 0.85 }}>Favorites</div>
                <div
                  style={{
                    maxHeight: 150,
                    minHeight: 150,
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
                    const bg = isSelected ? "rgba(0,120,212,0.12)" : (isHovered ? "rgba(0,120,212,0.10)" : "transparent");
                    return (
                      <div
                        key={id || `${name}_${i}`}
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
                        style={{ ...rowStyle, background: bg }}
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
              <div style={{ flex: "1 1 auto" }}>
                <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
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

                <div style={{ textAlign: "center", fontSize: 14, fontWeight: 600, marginTop: 18, opacity: 0.85, userSelect: "none" }}>
                  ⇄&nbsp;&nbsp;&nbsp;Double-click to transfer&nbsp;&nbsp;&nbsp;⇄
                </div>
              </div>
            </div>
          </div>
        </>
      )}
{activeTab === "Settings" && (
        <div style={{ fontSize: 13, opacity: 0.85 }}>Settings (coming soon)</div>
      )}
      {/* Global actions (outside tabs) */}
      <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 8, paddingTop: 8, borderTop: "1px solid #e0e0e0" }}>
        <button
          type="button"
          onClick={() => { try { if (window.Office?.context?.ui?.messageParent) { window.Office.context.ui.messageParent(JSON.stringify({ type: "cancel" })); } else { window.close?.(); } } catch (e) { console.error("Cancel failed:", e); window.close?.(); } }}
          style={{
            padding: "6px 14px",
            fontSize: 12,
            border: "1px solid #c8c8c8",
            borderRadius: 6,
            background: "#f5f5f5",
            cursor: "pointer",
          }}
        >
          Cancel
        </button>
      </div>
    </div>
  );
}


const addFavorite = (sheetId) => {
  if (!sheetId) return;
  // Until engine/storage Patch 2, use toggleFavorite to persist add/remove.
  onToggleFavorite(sheetId);
};

const removeFavorite = (sheetId) => {
  if (!sheetId) return;
  // Until engine/storage Patch 2, use toggleFavorite to persist add/remove.
  onToggleFavorite(sheetId);
};

const requestMoveFavorite = (sheetId, direction) => {
  if (!sheetId) return;
  if (direction !== "up" && direction !== "down") return;
  try {
    Office.context.ui.messageParent(JSON.stringify({ type: "moveFavorite", sheetId, direction }));
  } catch (err) {
    // It's OK if not supported yet.
    console.warn("messageParent(moveFavorite) failed:", err);
  }
};


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