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
  const s = Office.context.roamingSettings.get("settings") || {};
  return { ...DEFAULT_SETTINGS, ...s };
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
