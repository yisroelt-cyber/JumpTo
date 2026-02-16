// src/services/settingsTrace.js
// Diagnostics removed (no-op stubs retained to avoid churn in import graphs during incremental cleanup).
// If you reintroduce diagnostics later, restore the prior implementation from history.

/**
 * Diagnostics stub: previously appended settings snapshots to an ORTS-backed log.
 * @returns {Promise<void>}
 */
export function settingsTraceAppend() {
  return Promise.resolve();
}

/**
 * Diagnostics stub: previously flushed accumulated settings trace log to the diagnostics sheet.
 * @returns {Promise<void>}
 */
export function diagFlushSettingsTrace() {
  return Promise.resolve();
}

/**
 * Retained for compatibility with any ad-hoc scripts that referenced the key.
 */
export function settingsTraceKey() {
  return "JumpTo.Diag.SettingsTraceLog";
}
