import {
  officeReady,
  getSettings,
  saveSettings,
  rememberMailbox,
  readModeMailbox,
  ensureMasterCategory,
  currentItemCategories,
  openItem,
  getPending,
  savePending,
} from "./lib/office.js";
import { checkAll, trackSentEmail, findSentCopy, stopTracking, restIdFor, setKnownMailboxes } from "./lib/tracker.js";
import { parseHolidays } from "./lib/logic.js";
import { CATEGORY } from "./lib/config.js";

const $ = (id) => document.getElementById(id);
let settings;
let currentMbx = "me";

function setStatus(text, isError = false) {
  $("status").textContent = text;
  $("status").className = isError ? "error" : "muted";
}

function fmtDate(ymd) {
  if (!ymd) return "";
  const [y, m, d] = ymd.split("-").map(Number);
  return new Date(y, m - 1, d).toLocaleDateString(undefined, { weekday: "short", day: "numeric", month: "short" });
}

const names = (rs) => (rs || []).map((r) => r.emailAddress?.name || r.emailAddress?.address).join(", ");
const mbxLabel = (m) => (m === "me" ? "Your mailbox" : m);

// ---------------------------------------------------------------- current email card
async function refreshCurrent() {
  const item = Office.context.mailbox.item;
  if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message) {
    $("current").hidden = true;
    return;
  }
  $("current").hidden = false;
  $("currentSubject").textContent = item.subject || "(no subject)";
  currentMbx = await readModeMailbox();
  $("currentMailbox").textContent = mbxLabel(currentMbx);
  const tracked = (await currentItemCategories()).includes(CATEGORY);
  $("currentTracked").hidden = !tracked;
  $("currentUntracked").hidden = tracked;
}

async function onTrack() {
  const item = Office.context.mailbox.item;
  $("trackBtn").disabled = true;
  try {
    setStatus("Tracking…");
    if (currentMbx === "me") await ensureMasterCategory(CATEGORY);
    else settings = await rememberMailbox(currentMbx);
    setKnownMailboxes(settings.mailboxes);
    const state = await trackSentEmail(currentMbx, item.internetMessageId, Number($("trackDays").value), settings.holidays);
    setStatus(`Tracking. First follow-up due ${fmtDate(state.nextDue)}.`);
    $("currentTracked").hidden = false;
    $("currentUntracked").hidden = true;
    loadSettingsForm();
  } catch (e) {
    setStatus(e.message, true);
  } finally {
    $("trackBtn").disabled = false;
  }
}

async function onUntrack() {
  try {
    setStatus("Stopping…");
    setKnownMailboxes(settings.mailboxes);
    const msg = await findSentCopy(currentMbx, Office.context.mailbox.item.internetMessageId);
    if (!msg) throw new Error("Couldn't find the sent copy of this email.");
    await stopTracking(currentMbx, msg.id);
    setStatus("Tracking stopped.");
    $("currentTracked").hidden = true;
    $("currentUntracked").hidden = false;
  } catch (e) {
    setStatus(e.message, true);
  }
}

// ---------------------------------------------------------------- results list
const ORDER = { draftCreated: 0, draftWaiting: 1, error: 2, replied: 3, waiting: 4, ignored: 5 };

function renderResults(results) {
  const box = $("results");
  box.innerHTML = "";
  if (!results.length) {
    box.innerHTML = `<p class="muted pad">No tracked emails.</p>`;
    return;
  }
  const multi = settings.mailboxes.length > 1;
  results.sort((a, b) => ORDER[a.kind] - ORDER[b.kind]);
  for (const r of results) {
    const el = document.createElement("div");
    el.className = `card row ${r.kind}`;
    const line = {
      draftCreated: `Follow Up #${r.number} draft ready in Drafts`,
      draftWaiting: `Follow Up #${(r.state?.sent || 0) + 1} draft waiting – not sent yet`,
      replied: `Response received${r.by ? " from " + r.by : ""} – tracking stopped`,
      waiting: `Next check ${fmtDate(r.state?.nextDue)}${r.state?.sent ? ` · ${r.state.sent} sent` : ""}${r.state?.draftId ? " · draft pending" : ""}`,
      ignored: r.note,
      error: `Error: ${r.note}`,
    }[r.kind];
    el.innerHTML = `
      <div class="subject"></div>
      <div class="muted small to"></div>
      <div class="line"></div>
      ${r.bounced ? `<div class="error small">A delivery failure was received in this thread.</div>` : ""}
      <div class="actions"></div>`;
    el.querySelector(".subject").textContent = r.subject || "(no subject)";
    el.querySelector(".to").textContent = (multi ? mbxLabel(r.mbx) + " → " : "") + names(r.to);
    el.querySelector(".line").textContent = line;
    const actions = el.querySelector(".actions");
    if (r.draftId) actions.append(button("Open draft", () => openDraft(r.mbx, r.draftId, r.draftLink)));
    if (r.id && ["waiting", "draftCreated", "draftWaiting", "error"].includes(r.kind)) {
      actions.append(
        button("Stop tracking", async (b) => {
          b.disabled = true;
          await stopTracking(r.mbx, r.id);
          el.querySelector(".line").textContent = "Tracking stopped.";
          actions.innerHTML = "";
        })
      );
    }
    box.append(el);
  }
}

function button(label, fn) {
  const b = document.createElement("button");
  b.textContent = label;
  b.onclick = () => Promise.resolve(fn(b)).catch((e) => setStatus(e.message, true));
  return b;
}

async function openDraft(mbx, immutableId, webLink) {
  const m = await restIdFor(mbx, immutableId);
  openItem(m.id, m.webLink || webLink);
}

async function onCheck() {
  $("checkBtn").disabled = true;
  try {
    await ensureMasterCategory(CATEGORY);
    settings = getSettings();
    const { results, discovered, scanFrom, pending } = await checkAll(settings, settings.holidays, setStatus, getPending());
    settings.scanFrom = scanFrom;
    await saveSettings(settings);
    await savePending(pending);
    const drafts = results.filter((r) => r.kind === "draftCreated").length;
    setStatus(
      `Done. ${results.length} tracked · ${drafts} new follow-up draft${drafts === 1 ? "" : "s"}` +
        (discovered ? ` · ${discovered} newly picked up from Sent Items` : "")
    );
    renderResults(results);
  } catch (e) {
    setStatus(e.message, true);
  } finally {
    $("checkBtn").disabled = false;
  }
}

// ---------------------------------------------------------------- settings
function loadSettingsForm() {
  $("setTemplate").value = settings.template;
  $("setHolidays").value = (settings.holidays || []).join("\n");
  $("setDefaultDays").value = String(settings.defaultDays);
  $("setMailboxes").value = settings.mailboxes.filter((m) => m !== "me").join("\n");
}

async function onSaveSettings() {
  settings.template = $("setTemplate").value.trim() || settings.template;
  settings.holidays = parseHolidays($("setHolidays").value);
  settings.defaultDays = Number($("setDefaultDays").value);
  const shared = $("setMailboxes")
    .value.split(/[\s,;]+/)
    .map((m) => m.trim().toLowerCase())
    .filter((m) => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(m));
  settings.mailboxes = ["me", ...new Set(shared)];
  try {
    await saveSettings(settings);
    $("settingsStatus").textContent = "Saved (applies on both Macs).";
    loadSettingsForm();
  } catch (e) {
    $("settingsStatus").textContent = "Couldn't save: " + e.message;
  }
}

// ---------------------------------------------------------------- boot
officeReady().then(async () => {
  settings = getSettings();
  loadSettingsForm();
  $("checkBtn").onclick = onCheck;
  $("trackBtn").onclick = onTrack;
  $("untrackBtn").onclick = onUntrack;
  $("saveSettingsBtn").onclick = onSaveSettings;
  await refreshCurrent();
  // Pinned task pane: follow the selected email.
  if (Office.context.requirements.isSetSupported("Mailbox", "1.5")) {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => refreshCurrent());
  }
});
