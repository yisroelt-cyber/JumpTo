// src/services/rowHeightTrace.js
/* global Excel, OfficeRuntime */

// RowHeightPreset persistence tracing (diagnostics-only).
// Appends trace entries into OfficeRuntime.storage, periodically flushing to
// the workbook diagnostics sheet to avoid unbounded growth.

const ROW_HEIGHT_KEY = "JumpTo.Option.RowHeightPreset";
const TRACE_LOG_KEY = "JumpTo.Diag.RowHeightTraceLog";

// Markers chosen to be visually unique and searchable.
const BEGIN_MARK = "<<<JT_RH_TRACE_BEGIN>>>";
const END_MARK = "<<<JT_RH_TRACE_END>>>";

// Diagnostics sheet target (existing VeryHidden sheet).
const DIAG_SHEET_NAME = "_JumpToAddinSettings";
const DIAG_COL_RANGE = "C1:C500";

// Keep the in-ORTS log modest. If it grows past this, flush to sheet.
// (User requested: if max is hit, flush and clear, potentially multiple times.)
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

async function ortsGet(key) {
  if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage?.getItem) return "";
  try {
    const v = await OfficeRuntime.storage.getItem(key);
    return v == null ? "" : String(v);
  } catch (e) {
    return "";
  }
}

async function ortsSet(key, value) {
  if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage?.setItem) return;
  try {
    await OfficeRuntime.storage.setItem(key, String(value == null ? "" : value));
  } catch (e) {
    // ignore (diagnostics-only)
  }
}

async function writeChunkToDiagSheet(chunk) {
  if (!chunk) return;
  if (typeof Excel === "undefined" || !Excel.run) return;

  try {
    await Excel.run(async (context) => {
      const sh = context.workbook.worksheets.getItem(DIAG_SHEET_NAME);
      const col = sh.getRange(DIAG_COL_RANGE);
      col.load("values");
      await context.sync();

      const values = col.values || [];
      let row = 0;
      for (let i = 0; i < values.length; i++) {
        const v = values[i] && values[i][0];
        if (v === "" || v == null) {
          row = i;
          break;
        }
        row = i + 1;
      }
      if (row >= values.length) row = values.length - 1;

      sh.getRange(`C${row + 1}`).values = [[chunk]];
      await context.sync();
    });
  } catch (e) {
    try {
      console.error("[RowHeightTrace] Failed to write trace chunk to sheet", e);
    } catch (e2) {
      // ignore
    }
  }
}

async function flushTraceLogInternal(reasonTag) {
  const existing = await ortsGet(TRACE_LOG_KEY);
  if (!existing) return;

  const header = `${nowIso()} | FLUSH | ${String(reasonTag || "")}`;
  const chunk = `${header}\n${existing}`;
  await writeChunkToDiagSheet(chunk);
  await ortsSet(TRACE_LOG_KEY, "");
}

/**
 * Append a trace entry containing module/function and the current ORTS rowHeight value.
 *
 * @param {string} moduleName
 * @param {string} functionName
 * @param {string} [note] optional extra context
 */
export function diagTraceRowHeight(moduleName, functionName, note) {
  return enqueue(async () => {
    const rh = await ortsGet(ROW_HEIGHT_KEY);
    const entry = `${BEGIN_MARK}\n${nowIso()} | ${String(moduleName || "")} | ${String(functionName || "")} | rowHeight: ${rh}${note ? ` | ${String(note)}` : ""}\n${END_MARK}`;

    const existing = await ortsGet(TRACE_LOG_KEY);
    const next = existing ? `${existing}\n${entry}` : entry;
    await ortsSet(TRACE_LOG_KEY, next);

    if (next.length >= FLUSH_CHAR_THRESHOLD) {
      await flushTraceLogInternal("threshold");
    }
  });
}

/**
 * Force-flush the accumulated trace log to the diagnostics sheet and clear ORTS log.
 *
 * @param {string} moduleName
 * @param {string} functionName
 * @param {string} [reason]
 */
export function diagFlushRowHeightTrace(moduleName, functionName, reason) {
  return enqueue(async () => {
    const tag = `${String(moduleName || "")}::${String(functionName || "")}${reason ? `:${String(reason)}` : ""}`;
    await flushTraceLogInternal(tag);
  });
}
