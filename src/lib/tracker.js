// Follow-up tracking via Microsoft Graph, in your own mailbox and in shared
// mailboxes you have Full Access to (e.g. env-doc@gsk.com.sg).
//
// Source of truth lives on the sent email itself, so both Macs see the same state:
//   • category "Follow-up"            → email is being tracked (remove it in Outlook to stop)
//   • extended property FollowUpState → JSON {days, sent, nextDue, draftId, status}
// Emails marked in compose carry an Office.js custom property (fuDays) that is
// picked up from that mailbox's Sent Items on the next "Check follow-ups".
//
// A "mailbox" here is either "me" or the SMTP address of a shared mailbox.
import { graph, graphAll, q } from "./graph.js";
import { CATEGORY, CECP_PROP_ID, STATE_PROP_ID } from "./config.js";
import { toYMD, newState, parseState, decide, findResponse, isBounce, renderTemplate, matchPending, prunePending } from "./logic.js";

const MSG_FIELDS =
  "id,subject,sentDateTime,conversationId,toRecipients,ccRecipients,from,sender,isDraft,webLink,categories";

const root = (mbx) => (!mbx || mbx === "me" ? "/me" : `/users/${encodeURIComponent(mbx)}`);
const expandProp = (id) => `$expand=singleValueExtendedProperties($filter=id eq '${q(id)}')`;
const propValue = (msg, id) =>
  (msg.singleValueExtendedProperties || []).find((p) => p.id.toLowerCase() === id.toLowerCase())?.value;

// ---------------------------------------------------------------- identity
let meCache = null;
/** Your own addresses. "Mine" for a check = these + the mailbox being checked. */
export async function getMe() {
  if (meCache) return meCache;
  const me = await graph("GET", "/me?$select=displayName,mail,userPrincipalName,proxyAddresses");
  const addresses = new Set(
    [me.mail, me.userPrincipalName, ...(me.proxyAddresses || []).map((p) => p.replace(/^smtp:/i, ""))]
      .filter(Boolean)
      .map((a) => a.toLowerCase())
  );
  meCache = { name: me.displayName, addresses };
  return meCache;
}

// Shared mailboxes you work from. Messages from any of these (or from you) are
// never counted as a response, and count as "sent by you" for tracking.
let knownMailboxes = [];
export function setKnownMailboxes(list) {
  knownMailboxes = (list || []).filter((m) => m && m !== "me").map((m) => m.toLowerCase());
}

function ownAddresses(me, mbx) {
  const s = new Set([...me.addresses, ...knownMailboxes]);
  if (mbx && mbx !== "me") s.add(mbx.toLowerCase());
  return s;
}

const fromAddr = (m) => (m.from?.emailAddress?.address || "").toLowerCase();
const isOwn = (own, m) => own.has(fromAddr(m)) || own.has((m.sender?.emailAddress?.address || "").toLowerCase());

// ---------------------------------------------------------------- state writes
async function writeState(mbx, msg, state, { tracked }) {
  const cats = new Set(msg.categories || []);
  if (tracked) cats.add(CATEGORY);
  else cats.delete(CATEGORY);
  await graph("PATCH", `${root(mbx)}/messages/${msg.id}`, {
    categories: [...cats],
    singleValueExtendedProperties: [{ id: STATE_PROP_ID, value: JSON.stringify(state) }],
  });
}

export async function startTracking(mbx, msg, days, holidays) {
  const state = newState(days, msg.sentDateTime, holidays);
  await writeState(mbx, msg, state, { tracked: true });
  return state;
}

export async function stopTracking(mbx, msgId, status = "cancelled") {
  const msg = await graph("GET", `${root(mbx)}/messages/${msgId}?$select=id,categories&${expandProp(STATE_PROP_ID)}`);
  const state = parseState(propValue(msg, STATE_PROP_ID)) || { v: 1 };
  await writeState(mbx, msg, { ...state, status }, { tracked: false });
}

// ---------------------------------------------------------------- read-mode "Track this"
export async function findSentCopy(mbx, internetMessageId) {
  const own = ownAddresses(await getMe(), mbx);
  const hits = await graphAll(
    `${root(mbx)}/messages?$filter=internetMessageId eq '${q(internetMessageId)}'&$select=${MSG_FIELDS}&${expandProp(
      STATE_PROP_ID
    )}`,
    10
  );
  return hits.find((m) => !m.isDraft && isOwn(own, m)) || null;
}

export async function trackSentEmail(mbx, internetMessageId, days, holidays) {
  const msg = await findSentCopy(mbx, internetMessageId);
  if (!msg) throw new Error("Only emails sent from this mailbox can be tracked. Open it from Sent Items and try again.");
  return startTracking(mbx, msg, days, holidays);
}

// ---------------------------------------------------------------- discovery of compose-marked emails
async function discoverComposeMarked(mbx, since, holidays, pending = []) {
  if (pending.length) {
    const earliest = new Date(Math.min(...pending.map((p) => new Date(p.at).getTime())) - 5 * 60e3).toISOString();
    if (earliest < since) since = earliest;
  }
  const path =
    `${root(mbx)}/mailFolders/sentitems/messages?$filter=sentDateTime ge ${since}` +
    `&$select=${MSG_FIELDS}&${expandProp(CECP_PROP_ID)}&$top=100`;
  const sent = await graphAll(path, 2000);
  let found = 0;
  for (const m of sent) {
    const raw = propValue(m, CECP_PROP_ID);
    if (!raw) continue;
    let props;
    try {
      props = JSON.parse(raw);
    } catch {
      continue;
    }
    const days = Number(props.fuDays);
    if (!days || props.fuPromoted) continue;
    await startTracking(mbx, m, days, holidays);
    // Mark as picked up so a later scan doesn't restart tracking after a reply.
    const { fuDays, ...rest } = props;
    await graph("PATCH", `${root(mbx)}/messages/${m.id}`, {
      singleValueExtendedProperties: [{ id: CECP_PROP_ID, value: JSON.stringify({ ...rest, fuPromoted: "1" }) }],
    });
    found++;
  }

  // Pending marks from shared-mailbox compose (no marker on the email itself).
  const matchedKeys = [];
  if (pending.length) {
    const taken = new Set(sent.filter((m) => (m.categories || []).includes(CATEGORY) || propValue(m, CECP_PROP_ID)).map((m) => m.id));
    for (const [k, m] of matchPending(pending, sent, taken)) {
      const p = pending.find((x) => x.k === k);
      await startTracking(mbx, m, p.days, holidays);
      matchedKeys.push(k);
      found++;
    }
  }
  return { found, matchedKeys };
}

// ---------------------------------------------------------------- follow-up draft
function escapeHtml(s) {
  return s.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]);
}

async function createFollowUpDraft(mbx, original, number, template, own) {
  const notOwn = (r) => !own.has((r.emailAddress?.address || "").toLowerCase());
  let to = (original.toRecipients || []).filter(notOwn);
  let cc = (original.ccRecipients || []).filter(notOwn);
  if (!to.length && cc.length) [to, cc] = [cc, []];

  // Forward (not reply) so the original attachments come along. Stays in the same conversation.
  // Created in the same mailbox, so it is sent from that mailbox (e.g. env-doc@).
  const draft = await graph("POST", `${root(mbx)}/messages/${original.id}/createForward`, {
    message: { toRecipients: to, ccRecipients: cc },
  });

  const full = await graph("GET", `${root(mbx)}/messages/${draft.id}?$select=body,webLink`);
  const intro =
    `<div style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:11pt">` +
    escapeHtml(renderTemplate(template, number)).replace(/\n/g, "<br>") +
    `</div><br>`;
  let html = full.body?.content || "";
  html = /<body[^>]*>/i.test(html) ? html.replace(/<body[^>]*>/i, (m) => m + intro) : intro + html;

  // Clear categories so the sent follow-up itself is never treated as a new tracked email.
  await graph("PATCH", `${root(mbx)}/messages/${draft.id}`, {
    body: { contentType: "html", content: html },
    categories: [],
  });
  return { id: draft.id, webLink: full.webLink || draft.webLink };
}

// ---------------------------------------------------------------- the check
async function checkMailbox(mbx, settings, holidays, me, onProgress, pending) {
  const today = toYMD(new Date());
  const own = ownAddresses(me, mbx);
  const label = mbx === "me" ? "your mailbox" : mbx;

  onProgress(`${label}: looking for newly sent emails marked for follow-up…`);
  const since =
    (settings.scanFrom && settings.scanFrom[mbx]) ||
    new Date(Date.now() - settings.scanDays * 864e5).toISOString();
  const { found: discovered, matchedKeys } = await discoverComposeMarked(mbx, since, holidays, pending);

  onProgress(`${label}: loading tracked emails…`);
  const tracked = await graphAll(
    `${root(mbx)}/messages?$filter=categories/any(c:c eq '${q(CATEGORY)}')&$select=${MSG_FIELDS}&${expandProp(
      STATE_PROP_ID
    )}&$top=100`,
    500
  );

  const results = [];
  let i = 0;
  for (const msg of tracked) {
    i++;
    if (msg.isDraft) continue;
    onProgress(`${label}: checking ${i} of ${tracked.length}…`);
    const base = { mbx, id: msg.id, subject: msg.subject, to: msg.toRecipients, webLink: msg.webLink };
    try {
      if (!isOwn(own, msg)) {
        results.push({ ...base, kind: "ignored", note: "Not sent from this mailbox – remove the Follow-up category." });
        continue;
      }
      let state = parseState(propValue(msg, STATE_PROP_ID));
      // Category added by hand in Outlook, or re-added after a reply/cancel → (re)start.
      if (!state || state.status !== "active") {
        const startFrom = state ? new Date().toISOString() : msg.sentDateTime;
        state = await startTracking(mbx, { ...msg, sentDateTime: startFrom }, state?.days || settings.defaultDays, holidays);
      }

      const ctx = {};
      if (state.draftId) {
        const d = await graph(
          "GET",
          `${root(mbx)}/messages/${state.draftId}?$select=id,isDraft,sentDateTime,webLink`,
          undefined,
          { allow404: true }
        );
        ctx.draft = d ? { exists: true, isDraft: d.isDraft, sentDateTime: d.sentDateTime, webLink: d.webLink } : { exists: false };
      }

      const due = state.nextDue <= today || !!state.draftId;
      if (!due) {
        results.push({ ...base, kind: "waiting", state });
        continue;
      }

      const conv = await graphAll(
        `${root(mbx)}/messages?$filter=conversationId eq '${q(msg.conversationId)}'` +
          `&$select=id,from,subject,receivedDateTime,sentDateTime,isDraft&$top=50`,
        300
      );
      const response = findResponse(conv, msg.sentDateTime, own);
      ctx.responded = !!response;
      const bounced = conv.some((m) => isBounce(m) && new Date(m.receivedDateTime) > new Date(msg.sentDateTime));

      const r = decide(state, ctx, today, holidays);
      if (r.action === "stop") {
        await writeState(mbx, msg, r.state, { tracked: false });
        results.push({ ...base, kind: "replied", by: response?.from?.emailAddress?.name || response?.from?.emailAddress?.address });
      } else if (r.action === "waitDraft") {
        results.push({ ...base, kind: "draftWaiting", state, draftId: state.draftId, draftLink: ctx.draft?.webLink, bounced });
      } else if (r.action === "createDraft") {
        const draft = await createFollowUpDraft(mbx, msg, r.number, settings.template, own);
        const next = { ...r.state, draftId: draft.id };
        await writeState(mbx, msg, next, { tracked: true });
        results.push({ ...base, kind: "draftCreated", number: r.number, state: next, draftId: draft.id, draftLink: draft.webLink, bounced });
      } else {
        await writeState(mbx, msg, r.state, { tracked: true });
        results.push({ ...base, kind: "waiting", state: r.state });
      }
    } catch (e) {
      results.push({ ...base, kind: "error", note: e.message });
    }
  }
  return { results, discovered, matchedKeys };
}

/**
 * Runs on demand only ("Check follow-ups now"), across every mailbox in settings.mailboxes.
 *  - not yet due            → no mailbox query at all
 *  - due, response found    → tracking stops, category removed
 *  - due, no response       → Follow Up #N draft created in that mailbox, same conversation
 *  - draft already waiting  → reminds you to send it
 */
export async function checkAll(settings, holidays, onProgress = () => {}, pendingIn = []) {
  setKnownMailboxes(settings.mailboxes);
  const me = await getMe();
  const all = [];
  let discovered = 0;
  let pending = prunePending(pendingIn);
  const scanFrom = { ...(typeof settings.scanFrom === "object" ? settings.scanFrom : {}) };
  for (const mbx of settings.mailboxes) {
    try {
      const mine = pending.filter((p) => p.mbx === mbx || p.mbx === "*");
      const r = await checkMailbox(mbx, settings, holidays, me, onProgress, mine);
      all.push(...r.results);
      discovered += r.discovered;
      pending = pending.filter((p) => !r.matchedKeys.includes(p.k));
      // Overlap by 2 days so late-arriving Sent Items aren't missed. Roams to the other Mac.
      scanFrom[mbx] = new Date(Date.now() - 2 * 864e5).toISOString();
    } catch (e) {
      all.push({ mbx, subject: mbx === "me" ? "Your mailbox" : mbx, kind: "error", note: e.message });
    }
  }
  return { results: all, discovered, scanFrom, pending };
}

/** Plain REST id + webLink for opening an item with Office.js. */
export async function restIdFor(mbx, immutableId) {
  return graph("GET", `${root(mbx)}/messages/${immutableId}?$select=id,webLink`, undefined, { immutable: false });
}
