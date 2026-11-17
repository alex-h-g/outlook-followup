// Promise wrappers around the callback-style Office.js APIs we use.
import { DEFAULT_SETTINGS } from "./config.js";

export const officeReady = () => new Promise((resolve) => Office.onReady(resolve));

function wrap(fn) {
  return new Promise((resolve, reject) =>
    fn((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error)))
  );
}

// ---- roaming settings (stored in your mailbox, shared by both Macs) ----

export function getSettings() {
  const s = { ...DEFAULT_SETTINGS, ...(Office.context.roamingSettings.get("settings") || {}) };
  if (typeof s.scanFrom !== "object" || !s.scanFrom) s.scanFrom = {}; // v1.0 stored a single string
  s.mailboxes = [...new Set(["me", ...(s.mailboxes || [])].map((m) => (m === "me" ? m : m.toLowerCase())))];
  return s;
}

/** Add a shared mailbox to the list checked by "Check follow-ups now". */
export async function rememberMailbox(mbx) {
  const s = getSettings();
  if (s.mailboxes.includes(mbx)) return s;
  s.mailboxes.push(mbx);
  await saveSettings(s);
  return s;
}

const ownAddress = () => (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase();

/**
 * Mailbox of the email being read: "me", or the shared mailbox's address.
 * Uses getSharedPropertiesAsync (fails for items in your own mailbox).
 */
export async function readModeMailbox() {
  const item = Office.context.mailbox.item;
  if (!item || typeof item.getSharedPropertiesAsync !== "function") return "me";
  try {
    const sp = await wrap((cb) => item.getSharedPropertiesAsync(cb));
    const owner = (sp && sp.owner ? String(sp.owner) : "").toLowerCase();
    return owner && owner !== ownAddress() ? owner : "me";
  } catch {
    return "me";
  }
}

/** Mailbox an email is being composed from, based on its From address. */
export async function composeMailbox() {
  const item = Office.context.mailbox.item;
  try {
    const from = await wrap((cb) => item.from.getAsync(cb));
    const addr = (from && from.emailAddress ? from.emailAddress : "").toLowerCase();
    return addr && addr !== ownAddress() ? addr : "me";
  } catch {
    return "me";
  }
}

export function saveSettings(s) {
  Office.context.roamingSettings.set("settings", s);
  return wrap((cb) => Office.context.roamingSettings.saveAsync(cb));
}

// ---- compose-time marker (Office.js custom properties → Sent Items copy) ----
export async function getComposeDays() {
  const props = await wrap((cb) => Office.context.mailbox.item.loadCustomPropertiesAsync(cb));
  return Number(props.get("fuDays")) || 0;
}

export async function setComposeDays(days) {
  const item = Office.context.mailbox.item;
  const props = await wrap((cb) => item.loadCustomPropertiesAsync(cb));
  if (days) props.set("fuDays", String(days));
  else props.remove("fuDays");
  await wrap((cb) => props.saveAsync(cb));
  // Persist to the draft so the property travels with the item when sent.
  try {
    await wrap((cb) => item.saveAsync(cb));
  } catch {
    /* non-fatal: some clients save automatically */
  }
}

// ---- categories ----
export async function ensureMasterCategory(name) {
  try {
    const list = await wrap((cb) => Office.context.mailbox.masterCategories.getAsync(cb));
    if ((list || []).some((c) => c.displayName === name)) return;
    await wrap((cb) =>
      Office.context.mailbox.masterCategories.addAsync(
        [{ displayName: name, color: Office.MailboxEnums.CategoryColor.Preset1 }],
        cb
      )
    );
  } catch {
    /* non-fatal: category still works, just without a colour */
  }
}

export async function currentItemCategories() {
  const item = Office.context.mailbox.item;
  if (!item || !item.categories) return [];
  try {
    return ((await wrap((cb) => item.categories.getAsync(cb))) || []).map((c) => c.displayName);
  } catch {
    return [];
  }
}

export function openItem(restId, webLink) {
  try {
    const ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
    Office.context.mailbox.displayMessageForm(ewsId);
  } catch {
    if (webLink) window.open(webLink, "_blank");
  }
}

// ---- pending compose marks (shared mailboxes, where custom properties can't be saved) ----
const LS_KEY = "fu-pending";

function localPending() {
  try {
    return JSON.parse(localStorage.getItem(LS_KEY) || "[]");
  } catch {
    return [];
  }
}

/** All pending marks: roaming (both Macs) + any that only made it to this Mac. */
export function getPending() {
  const roam = getSettings().pending || [];
  const local = localPending();
  const seen = new Set(roam.map((p) => p.k));
  return [...roam, ...local.filter((p) => !seen.has(p.k))];
}

export async function savePending(list) {
  const s = getSettings();
  s.pending = list;
  try {
    await saveSettings(s);
    try {
      localStorage.removeItem(LS_KEY);
    } catch {}
  } catch {
    // Roaming settings unavailable in this context → keep on this Mac until the next check.
    localStorage.setItem(LS_KEY, JSON.stringify(list));
  }
}

/** What we can see of the draft right now. */
export async function composeSnapshot() {
  const item = Office.context.mailbox.item;
  const get = (fn) => wrap(fn).catch(() => null);
  const subject = (await get((cb) => item.subject.getAsync(cb))) || "";
  const to = ((await get((cb) => item.to.getAsync(cb))) || []).map((r) => (r.emailAddress || "").toLowerCase());
  return { subject, to, conv: item.conversationId || null };
}
