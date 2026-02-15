// Settings trace (diagnostics-only): captures dialog + commands settings snapshots into ORTS,
// and supports flushing to the diagnostics sheet (Column C) via the existing diag sheet pipeline.
//
// Design goals:
// - behavior neutral (no settings mutations here)
// - append-safe, bounded size
// - low risk of breaking Office.js runtime (defensive try/catch)
//
// NOTE: This module writes to OfficeRuntime.storage. In the dialog, messages are sent to the parent,
// and the parent writes entries so they can be flushed to the sheet.

const ORTS_KEY = "JumpTo.Diag.SettingsTraceLog";
const MAX_CHARS = 45000; // bounded; keep last ~45k chars

function nowIso() {
  try { return new Date().toISOString(); } catch (e) { return String(Date.now()); }
}

function safeJson(obj) {
  try { return JSON.stringify(obj); } catch (e) {
    try { return String(obj); } catch (e2) { return "[unserializable]"; }
  }
}

async function ortsGet(key) {
  try {
    if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage) return "";
    const v = await OfficeRuntime.storage.getItem(key);
    return typeof v === "string" ? v : (v == null ? "" : String(v));
  } catch (e) { return ""; }
}

async function ortsSet(key, value) {
  try {
    if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage) return;
    await OfficeRuntime.storage.setItem(key, String(value || ""));
  } catch (e) { /* ignore */ }
}

function trimToMax(s) {
  const str = String(s || "");
  if (str.length <= MAX_CHARS) return str;
  return str.slice(str.length - MAX_CHARS);
}

function formatEntry(moduleName, funcName, tag, snapshot, note) {
  const head = `${nowIso()} | ${String(moduleName || "unknown")} | ${String(funcName || "unknown")} | ${String(tag || "")}`;
  const snap = snapshot ? safeJson(snapshot) : "{}";
  const n = note ? safeJson(note) : "";
  return `${head} | snap:${snap}${n ? " | note:" + n : ""}\n`;
}

export async function settingsTraceAppend(moduleName, funcName, tag, snapshot, note) {
  const line = formatEntry(moduleName, funcName, tag, snapshot, note);
  try {
    const prev = await ortsGet(ORTS_KEY);
    const next = trimToMax(prev + line);
    await ortsSet(ORTS_KEY, next);
  } catch (e) { /* ignore */ }
}

export async function settingsTraceRead() { return await ortsGet(ORTS_KEY); }
export async function settingsTraceClear() { await ortsSet(ORTS_KEY, ""); }
export function settingsTraceKey() { return ORTS_KEY; }
