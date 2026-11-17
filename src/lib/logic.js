// Pure logic – no Office.js, no network. Unit-tested in test/logic.test.mjs.

/** Local calendar date as YYYY-MM-DD. */
export function toYMD(d) {
  const dt = d instanceof Date ? d : new Date(d);
  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, "0");
  const day = String(dt.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

/**
 * Add n business days (Mon–Fri, skipping holidays) to the local date of `from`.
 * Sent on Friday + 3 business days => Wednesday.
 */
export function addBusinessDays(from, n, holidays = []) {
  const hol = new Set(holidays);
  const dt = from instanceof Date ? new Date(from) : new Date(from);
  dt.setHours(12, 0, 0, 0); // avoid DST edge cases
  let added = 0;
  while (added < n) {
    dt.setDate(dt.getDate() + 1);
    const dow = dt.getDay();
    if (dow === 0 || dow === 6) continue;
    if (hol.has(toYMD(dt))) continue;
    added++;
  }
  return toYMD(dt);
}

export function isDue(state, today = toYMD(new Date())) {
  return state && state.status === "active" && state.nextDue <= today;
}

const AUTO_SUBJECT =
  /^(automatic reply|auto(matic)?[- ]?reply|autoreply|out of (the )?office|ooo\b|abwesenheit|r[ée]ponse automatique|undeliverable|undelivered|delivery (status notification|has failed)|mail delivery (failed|subsystem)|returned mail|read:|not read:|delivered:)/i;
const AUTO_SENDER = /^(mailer-daemon|postmaster|microsoftexchange[0-9a-f]*)@/i;

/** Out-of-office, bounces, read receipts – these never count as a response. */
export function isAutoMessage(msg) {
  const subj = (msg.subject || "").trim();
  const from = (msg.from?.emailAddress?.address || "").toLowerCase();
  return AUTO_SUBJECT.test(subj) || AUTO_SENDER.test(from);
}

export function isBounce(msg) {
  return /^(undeliverable|undelivered|delivery (status notification|has failed)|mail delivery|returned mail)/i.test(
    (msg.subject || "").trim()
  ) || AUTO_SENDER.test((msg.from?.emailAddress?.address || "").toLowerCase());
}

/**
 * A response = any non-draft message in the same conversation, received after
 * the tracked email was sent, from anyone other than me, that isn't automatic.
 */
export function findResponse(conversationMessages, sentDateTime, myAddresses) {
  const mine = new Set([...myAddresses].map((a) => a.toLowerCase()));
  const sentAt = new Date(sentDateTime).getTime();
  return (
    conversationMessages.find((m) => {
      if (m.isDraft) return false;
      const from = (m.from?.emailAddress?.address || "").toLowerCase();
      if (!from || mine.has(from)) return false;
      const t = new Date(m.receivedDateTime || m.sentDateTime || 0).getTime();
      if (t <= sentAt) return false;
      return !isAutoMessage(m);
    }) || null
  );
}

export function newState(days, sentDateTime, holidays) {
  return {
    v: 1,
    days,
    sent: 0, // follow-ups actually sent so far
    nextDue: addBusinessDays(new Date(sentDateTime), days, holidays),
    draftId: null, // pending follow-up draft (immutable Graph ID)
    status: "active", // active | replied | cancelled
    started: new Date(sentDateTime).toISOString(),
  };
}

export function parseState(json) {
  if (!json) return null;
  try {
    const s = JSON.parse(json);
    return s && s.v === 1 ? s : null;
  } catch {
    return null;
  }
}

/**
 * Decide what to do with a due, active tracked email.
 * @param state  current state
 * @param ctx    { responded: bool, draft: null | {exists:false} | {exists:true,isDraft:bool,sentDateTime} }
 * @returns { action, state } where action ∈ stop | waitDraft | createDraft | notDue
 */
export function decide(state, ctx, today, holidays = []) {
  let s = { ...state };

  // A previously-created follow-up draft that has since been sent: count it and
  // restart the clock from when it was sent.
  if (s.draftId && ctx.draft) {
    if (ctx.draft.exists && ctx.draft.isDraft) {
      if (ctx.responded) return { action: "stop", state: { ...s, status: "replied" } };
      return { action: "waitDraft", state: s };
    }
    if (ctx.draft.exists && !ctx.draft.isDraft) {
      s.sent += 1;
      s.nextDue = addBusinessDays(new Date(ctx.draft.sentDateTime), s.days, holidays);
    }
    // deleted without sending: forget it; a fresh draft is made below if due
    s.draftId = null;
  }

  if (ctx.responded) return { action: "stop", state: { ...s, status: "replied" } };
  if (s.nextDue > today) return { action: "notDue", state: s };
  return { action: "createDraft", state: s, number: s.sent + 1 };
}

export function renderTemplate(template, n) {
  return String(template).replaceAll("{n}", String(n));
}

export function parseHolidays(text) {
  return String(text || "")
    .split(/[\s,]+/)
    .map((t) => t.trim())
    .filter((t) => /^\d{4}-\d{2}-\d{2}$/.test(t));
}

// ---------------------------------------------------------------- pending compose marks
// Used where Outlook won't let the add-in tag a draft (shared mailboxes on Mac).
// A pending mark records what was known when "follow up" was chosen; the next
// check matches it to the sent copy in that mailbox's Sent Items.

export function normSubject(s) {
  return String(s || "")
    .replace(/^\s*((re|fw|fwd|aw|wg|tr)\s*:\s*)+/i, "")
    .trim()
    .toLowerCase();
}

const addrSet = (list) =>
  [...new Set((list || []).map((a) => String(a?.emailAddress?.address ?? a ?? "").toLowerCase()).filter(Boolean))]
    .sort()
    .join(";");

/**
 * @param pending  [{k, at, subject, to:[addr], conv}]
 * @param sent     Graph messages (id, sentDateTime, subject, toRecipients, conversationId)
 * @param taken    Set of message ids already tracked / claimed
 * @returns Map k → message
 */
export function matchPending(pending, sent, taken = new Set()) {
  const out = new Map();
  const claimed = new Set(taken);
  const ordered = [...pending].sort((a, b) => new Date(a.at) - new Date(b.at));
  for (const p of ordered) {
    const from = new Date(p.at).getTime() - 2 * 60e3; // clock skew
    const until = new Date(p.at).getTime() + 7 * 864e5;
    let best = null;
    let bestScore = 0;
    for (const m of sent) {
      if (claimed.has(m.id)) continue;
      const t = new Date(m.sentDateTime).getTime();
      if (t < from || t > until) continue;
      let score = 0;
      if (p.conv && m.conversationId === p.conv) score += 4;
      if (p.subject && normSubject(m.subject) === normSubject(p.subject)) score += 2;
      if (p.to?.length && addrSet(p.to) === addrSet(m.toRecipients)) score += 1;
      // Need more than a recipients-only match unless that's all we had.
      const minScore = p.subject || p.conv ? 2 : 1;
      if (score < minScore) continue;
      if (score > bestScore || (score === bestScore && t < new Date(best.sentDateTime).getTime())) {
        best = m;
        bestScore = score;
      }
    }
    if (best) {
      out.set(p.k, best);
      claimed.add(best.id);
    }
  }
  return out;
}

export function prunePending(pending, now = Date.now()) {
  return (pending || []).filter((p) => now - new Date(p.at).getTime() < 14 * 864e5);
}
