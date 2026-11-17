// Compose pane: No follow-up (default) / 3 / 4 / 5 business days.
//
// Own mailbox  → hidden marker saved on the draft (Office.js custom property).
// Shared mailbox (e.g. env-doc@) → Outlook for Mac refuses to save that marker
//   ("Internal protocol error -2147467259"), so instead a *pending mark* is stored
//   in the add-in's settings: subject + To + conversation + time. "Check follow-ups"
//   matches it to the sent copy in that mailbox's Sent Items.
import {
  officeReady,
  getComposeDays,
  setComposeDays,
  composeMailbox,
  rememberMailbox,
  getPending,
  savePending,
  composeSnapshot,
} from "./lib/office.js";

const $ = (id) => document.getElementById(id);
const status = (t, err) => {
  $("status").textContent = t;
  $("status").className = err ? "error" : "muted";
};
const label = (d) => (d ? `Follow-up set: ${d} business days.` : "Follow-up is off for this email.");

let mbx = "me";
let pendingKey = null; // this draft's pending mark (shared mailboxes)
let lastSnap = "";
let watcher = null;

const newKey = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);

async function upsertPending(days) {
  const snap = await composeSnapshot();
  let list = getPending().filter((p) => p.k !== pendingKey);
  if (days) {
    pendingKey = pendingKey || newKey();
    list.push({ k: pendingKey, mbx, days, at: new Date().toISOString(), ...snap });
    lastSnap = JSON.stringify(snap);
  } else {
    pendingKey = null;
  }
  await savePending(list);
  return snap;
}

// Keep subject/recipients current while the pane is open, so the match is reliable
// even if you chose the follow-up before typing the subject.
function watch() {
  clearInterval(watcher);
  if (!pendingKey) return;
  watcher = setInterval(async () => {
    const snap = await composeSnapshot();
    const s = JSON.stringify(snap);
    if (s === lastSnap) return;
    lastSnap = s;
    const list = getPending().map((p) => (p.k === pendingKey ? { ...p, ...snap } : p));
    await savePending(list);
  }, 4000);
}

async function choose(days) {
  status("Saving…");
  if (mbx === "me") {
    try {
      await setComposeDays(days);
      status(label(days));
      return;
    } catch {
      // Outlook wouldn't save the marker and we couldn't tell which mailbox this is:
      // fall back to a pending mark that every checked mailbox will look for.
      mbx = "*";
    }
  }
  if (mbx !== "*") await rememberMailbox(mbx);
  const snap = await upsertPending(days);
  watch();
  status(
    days
      ? `${label(days)}${snap.subject ? "" : " Tip: keep this panel open until you've typed the subject."}`
      : label(0)
  );
}

officeReady().then(async () => {
  const radios = [...document.querySelectorAll('input[name="days"]')];
  mbx = await composeMailbox();

  // Show the current choice.
  let days = 0;
  if (mbx === "me") {
    days = await getComposeDays().catch(() => 0);
  } else {
    const snap = await composeSnapshot();
    const mine = getPending().find(
      (p) => p.mbx === mbx && ((snap.conv && p.conv === snap.conv) || (snap.subject && p.subject === snap.subject))
    );
    if (mine) {
      pendingKey = mine.k;
      days = mine.days;
      lastSnap = JSON.stringify(snap);
      watch();
    }
  }
  radios.forEach((r) => (r.checked = Number(r.value) === days));
  status(label(days) + (mbx !== "me" ? ` (${mbx})` : ""));

  radios.forEach((r) =>
    r.addEventListener("change", () => choose(Number(r.value)).catch((e) => status("Couldn't save: " + e.message, true)))
  );
});
