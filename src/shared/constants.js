// Shared constants (single source of truth)
// Keep storage caps and UI constraints aligned to prevent silent truncation.
export const MAX_RECENTS = 20;
export const MAX_FAVORITES = 50;

// Premium feature flags (placeholder — wire to real license token when licensing is implemented)
// Set to true locally to test premium behaviour without a license.
export const PREMIUM_FREQ_BUMP = false;

// Settings registry — single source of truth for per-setting metadata.
//
// scope:
//   "workbook" — persisted in the hidden _JumpToAddinSettings sheet; writes fail in read-only workbooks.
//   "global"   — persisted in OfficeRuntime.storage; always writable regardless of workbook state.
//
// readOnlyDisable:
//   true  — the UI control for this setting must be disabled (and explained) when the workbook is read-only.
//   false — the setting writes to ORTS only; safe to leave interactive in read-only mode.
//
// When adding a new setting, you must declare it here. This ensures read-only UI handling
// is never an afterthought.
export const SETTINGS_REGISTRY = {
  // Workbook-scoped (writes to _JumpToAddinSettings sheet)
  oneDigitActivationEnabled: { scope: "workbook", readOnlyDisable: true },

  // Global-scoped (writes to OfficeRuntime.storage only)
  baselineOrder:       { scope: "global", readOnlyDisable: false },
  frequentOnTop:       { scope: "global", readOnlyDisable: false },
  favPercentManual:    { scope: "global", readOnlyDisable: false },
  recentsDisplayCount: { scope: "global", readOnlyDisable: false },
  rowHeightPreset:     { scope: "global", readOnlyDisable: false },
};

// Feature-level read-only registry — for UI features (not individual settings)
// that involve workbook writes and must be disabled or restricted in read-only mode.
export const FEATURES_REGISTRY = {
  favorites: { scope: "workbook", readOnlyDisable: true },
};
