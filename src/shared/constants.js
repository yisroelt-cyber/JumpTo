// 2026-04-09 12:00 PM EDT
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

// ─── Licensing ORTS keys ────────────────────────────────────────────────────
// All licensing state is stored in OfficeRuntime.storage (per-machine, per-user).

export const LIC_MACHINE_ID          = "JumpTo.Licensing.MachineId";         // locally generated GUID, set at activation
export const LIC_MACHINE_HASH        = "JumpTo.Licensing.MachineHash";       // trial: random GUID; post-activation: same or derived
export const LIC_LICENSE_KEY         = "JumpTo.Licensing.LicenseKey";        // post-activation only
export const LIC_LICENSE_STATUS      = "JumpTo.Licensing.LicenseStatus";     // "trial" | "retrial" | "active" | "expired" | "revoked" | "cancelled"
export const LIC_LICENSE_TYPE        = "JumpTo.Licensing.LicenseType";       // "individual" | "corporate" — only present when status is "revoked"
export const LIC_TIER                = "JumpTo.Licensing.Tier";              // "standard" | "premium"
export const LIC_LAST_CHECKIN        = "JumpTo.Licensing.LastCheckin";       // timestamp ms — MUST be read in existing getItems() batch
export const LIC_USE_DAYS_LOCAL      = "JumpTo.Licensing.UseDaysLocal";      // JSON array of "YYYY-MM-DD" strings, held until server confirms
export const LIC_TRIAL_ONSET         = "JumpTo.Licensing.TrialOnset";        // ISO date string, held until server confirms
export const LIC_TRIAL_ONSET_CONF    = "JumpTo.Licensing.TrialOnsetConfirmed";   // "true" once server confirms
export const LIC_WS_RANGE            = "JumpTo.Licensing.WorksheetRange";    // e.g. "11-20", held until server confirms
export const LIC_WS_RANGE_CONF       = "JumpTo.Licensing.WorksheetRangeConfirmed"; // "true" once server confirms
export const LIC_FRIENDLY_NAME       = "JumpTo.Licensing.FriendlyName";      // user-supplied machine name
export const LIC_MUJD_FAILURES       = "JumpTo.Licensing.MujdFailures";      // count of consecutive MUJD-confirmed server failures
export const LIC_USER_KEY_SOURCE     = "JumpTo.Licensing.UserKeySource";     // stable ID for UserKey derivation: license_id (individual) or employee_id (corporate)

// Worksheet survey — stored separately from licensing, synced on first check-in.
export const LIC_WS_SURVEY_DONE      = "JumpTo.Licensing.WsSurveyDone";      // "true" once user has answered and local data is set

// API base URL — will become https://api.leapsheet.com before launch.
export const API_BASE_URL = "https://leapsheet-worker.leapsheet.workers.dev";

// Checkin interval: fire if 3+ days (in ms) have passed since last check-in.
export const CHECKIN_INTERVAL_MS = 3 * 24 * 60 * 60 * 1000;

// MUJD: number of consecutive server-unreachable failures (under MUJD condition)
// before Licensed Standard users have Premium unlocked.
export const MUJD_UNLOCK_THRESHOLD = 5;
