// src/services/settingsTrace.js
/* global Excel, OfficeRuntime */

// Settings trace (diagnostics-only).
// Appends snapshot entries into OfficeRuntime.storage and can flush to the
// workbook diagnostics sheet (VeryHidden _JumpToAddinSettings, Column C).
//
// NOTE: Keep behavior-neutral: no settings mutations here.

const TRACE_LOG_KEY = "JumpTo.Diag.SettingsTraceLog";

// Markers chosen to be visually unique and searchable.
const BEGIN_MARK = "<<<JT_ST_TRACE_BEGIN>>>";
const END_MARK = "<<<JT_ST_TRACE_END>>>";

// Diagnostics sheet target (existing VeryHidden sheet).
const DIAG_SHEET_NAME = "_JumpToAddinSettings";
const DIAG_COL_RANGE = "C1:C500";

// Keep the in-ORTS log modest. If it grows past this, flush to sheet.
const FLUSH_CHAR_THRESHOLD = 14000;

// Serialize ORTS log mutations so we don't lose entries via interleaved
// read-append-write sequences.
let __appendChain = Promise.resolve();

function enqueue(op) {
  __appendChain = __appendChain.then(op, op);
  return __appendChain;
}

function nowIso() {
  try {
    return new Date().toISOString();
  } catch (e) {
    return String(Date.now());
  }
}

function safeJson(obj) {
  try {
    return JSON.stringify(obj);
  } catch (e) {
    try {
      return String(obj);
    } catch (e2) {
      return "[unserializable]";
    }
  }
}

async function ortsGet(key) {
  try {
    if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage) return "";
    const v = await OfficeRuntime.storage.getItem(key);
    return typeof v === "string" ? v : (v == null ? "" : String(v));
  } catch (e) {
    return "";
  }
}

async function ortsSet(key, value) {
  try {
    if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage) return;
    await OfficeRuntime.storage.setItem(key, String(value || ""));
  } catch (e) {
    // ignore
  }
}

async function writeChunkToDiagSheet(chunk) {
  try {
    if (typeof Excel === "undefined" || !Excel.run) return;

    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getItem(DIAG_SHEET_NAME);
      const range = ws.getRange(DIAG_COL_RANGE);
      range.load("values");
      await context.sync();

      const values = Array.isArray(range.values) ? range.values : [];
      const flat = values.map((r) => (Array.isArray(r) ? (r[0] == null ? "" : String(r[0])) : ""));

      // Find first empty cell; if none, overwrite from top.
      let idx = flat.findIndex((x) => !x);
      if (idx < 0) idx = 0;

      flat[idx] = String(chunk || "");
      range.values = flat.map((x) => [x]);
      await context.sync();
    });
  } catch (e) {
    // ignore
  }
}

async function flushTraceLogInternal(reasonTag) {
  const existing = await ortsGet(TRACE_LOG_KEY);

  const header = `${nowIso()} | FLUSH | ${String(reasonTag || "")}`;
  const chunk = `${header}\n${existing ? existing : "(empty settings trace log)"}`;

  await writeChunkToDiagSheet(chunk);

  // Always clear after a flush attempt so each run is isolated.
  await ortsSet(TRACE_LOG_KEY, "");
}

/**
 * Append a settings trace entry.
 *
 * @param {string} moduleName
 * @param {string} functionName
 * @param {string} tag
 * @param {object} snapshot
 * @param {object} note
 */
export function settingsTraceAppend(moduleName, functionName, tag, snapshot, note) {
  return enqueue(async () => {
    const entry = `${BEGIN_MARK}\n${nowIso()} | ${String(moduleName || "")} | ${String(functionName || "")} | ${String(tag || "")} | snap:${safeJson(snapshot || {})}${note ? ` | note:${safeJson(note)}` : ""}\n${END_MARK}`;

    const existing = await ortsGet(TRACE_LOG_KEY);
    const next = existing ? `${existing}\n${entry}` : entry;
    await ortsSet(TRACE_LOG_KEY, next);

    if (next.length >= FLUSH_CHAR_THRESHOLD) {
      await flushTraceLogInternal("threshold");
    }
  });
}

/**
 * Force-flush the accumulated settings trace log to the diagnostics sheet and clear ORTS log.
 */
export function diagFlushSettingsTrace(moduleName, functionName, reason) {
  return enqueue(async () => {
    const tag = `${String(moduleName || "")}::${String(functionName || "")}${reason ? `:${String(reason)}` : ""}`;
    await flushTraceLogInternal(tag);
  });
}


export function settingsTraceKey() { return TRACE_LOG_KEY; }
