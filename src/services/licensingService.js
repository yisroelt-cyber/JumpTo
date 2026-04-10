// 2026-04-10 12:00 PM EDT
/**
 * licensingService.js
 *
 * All licensing logic for LeapSheet Compatible (Office.js).
 * Runs in the commands.js (host/shared runtime) context.
 *
 * Responsibilities:
 *   - ORTS read/write for all licensing state
 *   - Use-day tracking (local accumulation, server sync on check-in)
 *   - Trial onset tracking
 *   - Activation (Endpoint 1)
 *   - Deregistration (Endpoint 3) — separate from activation
 *   - Check-in (Endpoint 2), including MUJD detection
 *   - UserKey derivation (SHA-256 of license_id or employee_id via SubtleCrypto)
 *   - machine_id and machine_hash generation
 *
 * Key constraints:
 *   - The LIC_LAST_CHECKIN timestamp is read as part of the EXISTING ORTS
 *     getItems() batch in buildDialogState() in commands.js — no separate call.
 *   - Check-in fires async in the background; does NOT block dialog open.
 *   - Check-in consequences surface at the NEXT dialog open.
 */

/* global OfficeRuntime */

import {
  LIC_MACHINE_ID,
  LIC_MACHINE_HASH,
  LIC_LICENSE_KEY,
  LIC_LICENSE_STATUS,
  LIC_LICENSE_TYPE,
  LIC_TIER,
  LIC_LAST_CHECKIN,
  LIC_USE_DAYS_LOCAL,
  LIC_TRIAL_ONSET,
  LIC_TRIAL_ONSET_CONF,
  LIC_WS_RANGE,
  LIC_WS_RANGE_CONF,
  LIC_FRIENDLY_NAME,
  LIC_MUJD_FAILURES,
  LIC_WS_SURVEY_DONE,
  LIC_USER_KEY_SOURCE,
  LIC_MACHINE_STATUS,
  LIC_RETRIAL_AVAILABLE,
  LIC_TAMPERED,
  LIC_UPGRADE_IN_PROGRESS,
  API_BASE_URL,
  CHECKIN_INTERVAL_MS,
  MUJD_UNLOCK_THRESHOLD,
} from "../shared/constants";

// ─── Utilities ───────────────────────────────────────────────────────────────

function generateGuid() {
  // RFC 4122 v4 UUID using crypto.getRandomValues when available.
  try {
    if (typeof crypto !== "undefined" && crypto.randomUUID) {
      return crypto.randomUUID();
    }
    // Fallback: manual construction.
    const buf = new Uint8Array(16);
    if (typeof crypto !== "undefined" && crypto.getRandomValues) {
      crypto.getRandomValues(buf);
    } else {
      for (let i = 0; i < 16; i++) buf[i] = Math.floor(Math.random() * 256);
    }
    buf[6] = (buf[6] & 0x0f) | 0x40;
    buf[8] = (buf[8] & 0x3f) | 0x80;
    const hex = Array.from(buf).map((b) => b.toString(16).padStart(2, "0")).join("");
    return `${hex.slice(0,8)}-${hex.slice(8,12)}-${hex.slice(12,16)}-${hex.slice(16,20)}-${hex.slice(20)}`;
  } catch (e) {
    // Last-resort fallback.
    return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
      const r = Math.random() * 16 | 0;
      return (c === "x" ? r : (r & 0x3 | 0x8)).toString(16);
    });
  }
}

async function sha256Hex(str) {
  try {
    const encoder = new TextEncoder();
    const data = encoder.encode(str);
    const hashBuffer = await crypto.subtle.digest("SHA-256", data);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map((b) => b.toString(16).padStart(2, "0")).join("");
  } catch (e) {
    // SubtleCrypto unavailable (non-secure context in some Office webviews).
    // Fall back to a simple deterministic hash — not cryptographic, but sufficient
    // as a stable identifier in this context.
    let h = 0;
    for (let i = 0; i < str.length; i++) {
      h = (Math.imul(31, h) + str.charCodeAt(i)) | 0;
    }
    return Math.abs(h).toString(16).padStart(8, "0").repeat(8).slice(0, 64);
  }
}

function todayIso() {
  return new Date().toISOString().slice(0, 10); // "YYYY-MM-DD"
}

function ortsGet(key) {
  if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage?.getItem) {
    return Promise.resolve(null);
  }
  return OfficeRuntime.storage.getItem(key).catch(() => null);
}

function ortsSet(key, value) {
  if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage?.setItem) {
    return Promise.resolve(null);
  }
  return OfficeRuntime.storage.setItem(key, value).catch(() => null);
}

function ortsGetMany(keys) {
  if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage?.getItems) {
    return Promise.resolve({});
  }
  return OfficeRuntime.storage.getItems(keys).catch(() => ({}));
}

function ortsSetMany(pairs) {
  // pairs: { [key]: value }
  if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage?.setItem) {
    return Promise.resolve();
  }
  return Promise.all(
    Object.entries(pairs).map(([k, v]) => OfficeRuntime.storage.setItem(k, v).catch(() => null))
  );
}

async function apiFetch(path, body) {
  const url = `${API_BASE_URL}${path}`;
  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  return resp.json();
}

// ─── MUJD detection ──────────────────────────────────────────────────────────

/**
 * Checks if the internet is generally reachable (via microsoft.com HEAD request).
 * Used only when our server is unreachable (MUJD check).
 */
async function isMicrosoftReachable() {
  try {
    await fetch("https://www.microsoft.com/favicon.ico", {
      method: "HEAD",
      cache: "no-store",
      mode: "no-cors",
    });
    // no-cors: resp.ok is always false, but no throw = reachable
    return true;
  } catch (e) {
    return false;
  }
}

// ─── Machine identity ─────────────────────────────────────────────────────────

/**
 * Returns the machine_hash for this machine.
 * Trial: randomly generated GUID stored in ORTS on first launch.
 * (Post-activation machine_id is a separate value set at activation time.)
 */
export async function ensureMachineHash() {
  let hash = await ortsGet(LIC_MACHINE_HASH);
  if (!hash) {
    hash = generateGuid();
    await ortsSet(LIC_MACHINE_HASH, hash);
  }
  return hash;
}

/**
 * Returns the post-activation machine_id. Null if not yet activated.
 */
export async function getMachineId() {
  return ortsGet(LIC_MACHINE_ID);
}

// ─── Use-day tracking ────────────────────────────────────────────────────────

/**
 * Records today as a use-day if not already recorded.
 * Called once per "application load" (first dialog open of the session).
 * Stores locally; synced to server on next check-in.
 */
export async function recordUseDay() {
  try {
    const raw = await ortsGet(LIC_USE_DAYS_LOCAL);
    const days = raw ? JSON.parse(raw) : [];
    const today = todayIso();
    if (Array.isArray(days) && days.includes(today)) return; // already recorded
    const next = Array.isArray(days) ? [...days, today] : [today];
    await ortsSet(LIC_USE_DAYS_LOCAL, JSON.stringify(next));
  } catch (e) {
    // ignore
  }
}

// ─── Full licensing state read ────────────────────────────────────────────────

/**
 * Reads all licensing state from ORTS in one batch.
 * Returns a normalized object used by buildDialogState() in commands.js.
 *
 * NOTE: LIC_LAST_CHECKIN is ALREADY read in the existing getItems() batch
 * in commands.js. This function reads the rest of the licensing keys.
 * The caller (commands.js) merges last_checkin from the main batch into
 * the result of this function.
 */
export async function readLicensingState() {
  const keys = [
    LIC_MACHINE_ID,
    LIC_MACHINE_HASH,
    LIC_LICENSE_KEY,
    LIC_LICENSE_STATUS,
    LIC_LICENSE_TYPE,
    LIC_TIER,
    LIC_USE_DAYS_LOCAL,
    LIC_TRIAL_ONSET,
    LIC_TRIAL_ONSET_CONF,
    LIC_WS_RANGE,
    LIC_WS_RANGE_CONF,
    LIC_FRIENDLY_NAME,
    LIC_MUJD_FAILURES,
    LIC_WS_SURVEY_DONE,
    LIC_USER_KEY_SOURCE,
    LIC_MACHINE_STATUS,
    LIC_RETRIAL_AVAILABLE,
    LIC_TAMPERED,
    LIC_UPGRADE_IN_PROGRESS,
  ];
  const raw = await ortsGetMany(keys);

  const licenseStatus = raw[LIC_LICENSE_STATUS] || "trial";
  const tier = raw[LIC_TIER] || "standard";
  const mujdFailures = parseInt(raw[LIC_MUJD_FAILURES] || "0", 10) || 0;

  return {
    machine_id:              raw[LIC_MACHINE_ID]     || null,
    machine_hash:            raw[LIC_MACHINE_HASH]   || null,
    license_key:             raw[LIC_LICENSE_KEY]    || null,
    license_status:          licenseStatus,
    license_type:            raw[LIC_LICENSE_TYPE]   || null,
    tier,
    use_days_local:          raw[LIC_USE_DAYS_LOCAL] ? (() => { try { return JSON.parse(raw[LIC_USE_DAYS_LOCAL]); } catch (e) { return []; } })() : [],
    trial_onset:             raw[LIC_TRIAL_ONSET]    || null,
    trial_onset_confirmed:   raw[LIC_TRIAL_ONSET_CONF] === "true",
    worksheet_range:         raw[LIC_WS_RANGE]       || null,
    worksheet_range_confirmed: raw[LIC_WS_RANGE_CONF] === "true",
    friendly_name:           raw[LIC_FRIENDLY_NAME]  || null,
    mujd_failures:           mujdFailures,
    ws_survey_done:          raw[LIC_WS_SURVEY_DONE] === "true",
    user_key_source:         raw[LIC_USER_KEY_SOURCE] || null,
    machine_status:          raw[LIC_MACHINE_STATUS]  || null,   // "unregistered" or null
    retrial_available:       raw[LIC_RETRIAL_AVAILABLE] === "true",
    tampered:                raw[LIC_TAMPERED] === "true",
    upgrade_in_progress:     raw[LIC_UPGRADE_IN_PROGRESS] === "true",
    // last_checkin injected by commands.js from main ORTS batch
  };
}

// ─── Trial onset initialization ───────────────────────────────────────────────

/**
 * Ensures trial onset date is recorded in ORTS on first launch.
 * No-op if already set.
 */
export async function ensureTrialOnset() {
  try {
    const existing = await ortsGet(LIC_TRIAL_ONSET);
    if (existing) return;
    await ortsSet(LIC_TRIAL_ONSET, todayIso());
  } catch (e) {
    // ignore
  }
}

// ─── Worksheet survey ─────────────────────────────────────────────────────────

/**
 * Stores the worksheet survey answer in ORTS.
 * Called when the user submits the mandatory survey.
 */
export async function saveWorksheetSurveyAnswer(range) {
  try {
    await ortsSetMany({
      [LIC_WS_RANGE]:      range,
      [LIC_WS_SURVEY_DONE]: "true",
      [LIC_WS_RANGE_CONF]:  "false", // not yet confirmed by server
    });
  } catch (e) {
    // ignore
  }
}

// ─── Check-in ─────────────────────────────────────────────────────────────────

/**
 * Performs a background check-in if 3+ days have passed since the last one.
 *
 * @param {object} licState  - result of readLicensingState() merged with last_checkin
 * @returns {Promise<void>}  - always resolves; errors are swallowed
 */
export async function maybeCheckin(licState) {
  try {
    const now = Date.now();
    const lastCheckin = licState.last_checkin ? parseInt(licState.last_checkin, 10) : 0;
    if (now - lastCheckin < CHECKIN_INTERVAL_MS) return;

    await performCheckin(licState);
  } catch (e) {
    // Checkin errors must never surface to the user.
  }
}

async function performCheckin(licState) {
  const now = Date.now();

  // Ensure machine_hash exists (creates on first call).
  const machine_hash = licState.machine_hash || (await ensureMachineHash());
  const machine_id   = licState.machine_id;

  let body;
  if (machine_id && licState.license_status === "active") {
    // Activated machine check-in.
    body = { machine_id };
  } else {
    // Trial machine check-in.
    const use_days = Array.isArray(licState.use_days_local) ? licState.use_days_local : [];
    body = {
      machine_hash,
      new_use_days: use_days,
    };

    // Piggyback trial_onset on first check-in (until server confirms).
    if (licState.trial_onset && !licState.trial_onset_confirmed) {
      body.trial_onset = { product: "compatible", date: licState.trial_onset };
    }

    // Piggyback worksheet_range on first check-in (until server confirms).
    if (licState.worksheet_range && !licState.worksheet_range_confirmed) {
      body.worksheet_range = licState.worksheet_range;
    }
  }

  let data;
  let serverReachable = true;
  try {
    data = await apiFetch("/checkin", body);
  } catch (e) {
    serverReachable = false;
  }

  if (!serverReachable) {
    // Server unreachable — run MUJD detection.
    await handleServerUnreachable(licState);
    return;
  }

  // Server reached — reset MUJD counter.
  await ortsSet(LIC_MUJD_FAILURES, "0");

  if (!data || data.result !== "ok") return;

  // Update last_checkin timestamp.
  await ortsSet(LIC_LAST_CHECKIN, String(now));

  // Process confirmed fields.
  const updates = {};

  if (data.trial_onset_confirmed) {
    updates[LIC_TRIAL_ONSET_CONF] = "true";
  }

  if (data.worksheet_range_confirmed) {
    updates[LIC_WS_RANGE_CONF] = "true";
    // Clear local use_days only after server confirms receipt.
    updates[LIC_USE_DAYS_LOCAL] = JSON.stringify([]);
  }

  // Update license_status if server returns a new state.
  if (data.license_status) {
    updates[LIC_LICENSE_STATUS] = data.license_status;
  }

  // Update tier if returned.
  if (data.tier) {
    updates[LIC_TIER] = data.tier;
  }

  // Update license_type (for revoked corporate message).
  if (data.license_type) {
    updates[LIC_LICENSE_TYPE] = data.license_type;
  } else if (data.license_status && data.license_status !== "revoked") {
    // Clear license_type when not revoked.
    updates[LIC_LICENSE_TYPE] = "";
  }

  // machine_status: "unregistered" when machine is not entitled for this license.
  // Absent means machine is in good standing — clear any stale value.
  if (data.machine_status === "unregistered") {
    updates[LIC_MACHINE_STATUS] = "unregistered";
  } else {
    updates[LIC_MACHINE_STATUS] = "";
  }

  // tampered: true when server use-day count exceeds client-reported count.
  // Accompanies trial or retrial status only. Clear when not present.
  updates[LIC_TAMPERED] = data.tampered === true ? "true" : "";

  // retrial_available: returned when license_status is "expired".
  if (data.retrial_available === true) {
    updates[LIC_RETRIAL_AVAILABLE] = "true";
  } else if (data.license_status && data.license_status !== "expired") {
    // Not expired — clear any stale retrial flag.
    updates[LIC_RETRIAL_AVAILABLE] = "";
  }

  if (Object.keys(updates).length > 0) {
    await ortsSetMany(updates);
  }
}

async function handleServerUnreachable(licState) {
  try {
    const microsoftReachable = await isMicrosoftReachable();
    if (!microsoftReachable) {
      // Plain offline — no MUJD. Don't increment counter.
      return;
    }

    // MUJD confirmed (internet up, our server down).
    const currentFailures = licState.mujd_failures || 0;
    const newFailures = currentFailures + 1;
    await ortsSet(LIC_MUJD_FAILURES, String(newFailures));
  } catch (e) {
    // ignore
  }
}

// ─── Effective license state (applying connection state overrides) ─────────────

/**
 * Computes the effective licensing state to expose to the dialog UI.
 * Applies state:open (MUJD) overrides as per LPD §15.
 *
 * state:open overrides (mujdActive):
 *   - trial / retrial / expired → treated as active (extends indefinitely)
 *   - active → continues normally; Standard tier unlocked to Premium
 *   - revoked / cancelled → NOT overridden (deliberate states)
 *   - machine_status: "unregistered" → NOT overridden (deliberate state)
 *
 * tampered flag:
 *   - When tampered=true and status is trial or retrial, falls back to
 *     trial_onset_date + 30 calendar days (handled by the dialog; passed through here).
 *   - state:open (mujdActive) takes priority over tampered.
 *
 * @param {object} licState - raw licensing state from readLicensingState()
 * @returns {object} - effective state for dialog consumption
 */
export function computeEffectiveLicenseState(licState) {
  const mujdActive = (licState.mujd_failures || 0) >= MUJD_UNLOCK_THRESHOLD;

  let effectiveStatus = licState.license_status || "trial";
  let mujdOverride = false;

  if (mujdActive) {
    if (
      effectiveStatus === "trial" ||
      effectiveStatus === "retrial" ||
      effectiveStatus === "expired"
    ) {
      effectiveStatus = "active"; // state:open: extend trial/retrial/expired
      mujdOverride = true;
    }
    // revoked and cancelled are NOT overridden — deliberate states.
  }

  // machine_status: "unregistered" forces restricted regardless of connection state.
  // This is a machine property, not a license property, and is never overridden by state:open.
  const machineUnregistered = licState.machine_status === "unregistered";

  // Restricted = any state that locks the UI to About tab only.
  const isRestricted =
    machineUnregistered ||
    effectiveStatus === "expired" ||
    effectiveStatus === "revoked" ||
    effectiveStatus === "cancelled";

  // Under state:open, Standard licensed users get Premium unlocked.
  const effectiveTier =
    mujdActive && !mujdOverride && effectiveStatus === "active"
      ? "premium"
      : (licState.tier || "standard");

  return {
    ...licState,
    effective_status:    effectiveStatus,
    effective_tier:      effectiveTier,
    is_restricted:       isRestricted,
    machine_unregistered: machineUnregistered,
    mujd_active:         mujdActive,
    mujd_override:       mujdOverride,
  };
}

// ─── Activation ───────────────────────────────────────────────────────────────

/**
 * Activates a license key on this machine (Endpoint 1).
 * If machineToDeregister is provided, routes to deregisterAndActivate (Endpoint 3) instead.
 * On success, stores the stable ID (license_id or employee_id) returned by
 * the server as LIC_USER_KEY_SOURCE for UserKey derivation.
 *
 * @param {string} licenseKey
 * @param {string} friendlyName
 * @param {string|null} machineToDeregister - machine_id to deregister (slots_full flow)
 * @returns {Promise<object>} - { status, tier, userKey } on activated; { status, machines } on slots_full; { status } otherwise
 */
export async function activateLicense(licenseKey, friendlyName, machineToDeregister) {
  // Deregistration flow: route to Endpoint 3.
  if (machineToDeregister) {
    return deregisterAndActivate(licenseKey, machineToDeregister, friendlyName);
  }

  // Ensure machine_hash and machine_id are ready.
  const machine_hash = await ensureMachineHash();

  // Generate a new locally-created machine_id for the post-activation identity.
  let machine_id = await ortsGet(LIC_MACHINE_ID);
  if (!machine_id) {
    machine_id = generateGuid();
    await ortsSet(LIC_MACHINE_ID, machine_id);
  }

  const body = {
    license_key:   licenseKey,
    machine_hash,
    machine_id,
    friendly_name: friendlyName || "",
  };

  const data = await apiFetch("/activate", body);

  if (data.result !== "ok") {
    throw new Error("Unexpected server response");
  }

  if (data.activation_status === "activated") {
    // Stable ID for UserKey derivation — license_id (individual) or employee_id (corporate).
    const userKeySource = data.license_id || data.employee_id || null;
    const userKey = userKeySource ? await sha256Hex(userKeySource) : null;

    // Persist activated state; clear machine_status (machine is now registered).
    await ortsSetMany({
      [LIC_LICENSE_KEY]:    licenseKey,
      [LIC_LICENSE_STATUS]: "active",
      [LIC_TIER]:           data.tier || "standard",
      [LIC_FRIENDLY_NAME]:  friendlyName || "",
      [LIC_MUJD_FAILURES]:  "0",
      [LIC_LICENSE_TYPE]:   "",
      [LIC_USER_KEY_SOURCE]: userKeySource || "",
      [LIC_MACHINE_STATUS]: "",
    });

    return { status: "activated", tier: data.tier || "standard", userKey };
  }

  if (data.activation_status === "slots_full") {
    return { status: "slots_full", machines: data.machines || [] };
  }

  if (data.activation_status === "invalid_key") {
    return { status: "invalid_key" };
  }

  if (data.activation_status === "rate_limited") {
    return { status: "rate_limited" };
  }

  return { status: "unknown", raw: data };
}

// ─── Deregistration ───────────────────────────────────────────────────────────

/**
 * Deregisters an existing machine and activates this one in its place (Endpoint 3).
 * Called when activation returns slots_full and the user selects a machine to remove.
 *
 * @param {string} licenseKey
 * @param {string} machineToDeregister  - machine_id of the machine to remove
 * @param {string} friendlyName
 * @returns {Promise<object>} - { status, tier } on activated; { status, nextSwitchAllowed } on rate_limited
 */
export async function deregisterAndActivate(licenseKey, machineToDeregister, friendlyName) {
  const machine_hash = await ensureMachineHash();
  let machine_id = await ortsGet(LIC_MACHINE_ID);
  if (!machine_id) {
    machine_id = generateGuid();
    await ortsSet(LIC_MACHINE_ID, machine_id);
  }

  const body = {
    license_key:           licenseKey,
    machine_to_deregister: machineToDeregister,
    new_machine_hash:      machine_hash,
    new_machine_id:        machine_id,
    friendly_name:         friendlyName || "",
  };

  const data = await apiFetch("/deregister", body);

  if (data.result !== "ok") {
    throw new Error("Unexpected server response");
  }

  if (data.switch_status === "activated") {
    // Stable ID returned directly by the deregister response.
    const userKeySource = data.license_id || data.employee_id || null;
    const userKey = userKeySource ? await sha256Hex(userKeySource) : null;

    await ortsSetMany({
      [LIC_LICENSE_KEY]:    licenseKey,
      [LIC_LICENSE_STATUS]: "active",
      [LIC_TIER]:           data.tier || "standard",
      [LIC_FRIENDLY_NAME]:  friendlyName || "",
      [LIC_MUJD_FAILURES]:  "0",
      [LIC_LICENSE_TYPE]:   "",
      [LIC_USER_KEY_SOURCE]: userKeySource || "",
      [LIC_MACHINE_STATUS]: "",
    });

    return { status: "activated", tier: data.tier || "standard", userKey };
  }

  if (data.switch_status === "rate_limited") {
    return { status: "rate_limited", nextSwitchAllowed: data.next_switch_allowed || null };
  }

  return { status: "unknown", raw: data };
}

// ─── UserKey ──────────────────────────────────────────────────────────────────

/**
 * Returns the UserKey for the current machine.
 * Post-activation: SHA-256 of LIC_USER_KEY_SOURCE (license_id or employee_id).
 * Trial: machine_hash.
 */
export async function getUserKey() {
  try {
    const userKeySource = await ortsGet(LIC_USER_KEY_SOURCE);
    if (userKeySource) {
      return sha256Hex(userKeySource);
    }
    // Trial fallback: use machine_hash as UserKey.
    const hash = await ortsGet(LIC_MACHINE_HASH);
    return hash || null;
  } catch (e) {
    return null;
  }
}
