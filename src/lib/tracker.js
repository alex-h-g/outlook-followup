// Follow-up tracking against your own mailbox via Microsoft Graph.
//
// Source of truth lives on the sent email itself, so both Macs see the same state:
//   • category "Follow-up"          → email is being tracked (remove it in Outlook to stop)
//   • extended property FollowUpState → JSON {days, sent, nextDue, draftId, status}
// Emails marked in compose carry an Office.js custom property (fuDays) that is
// picked up from Sent Items on the next "Check follow-ups".
import { graph, graphAll, q } from "./graph.js";
import { CATEGORY, CECP_PROP_ID, STATE_PROP_ID } from "./config.js";
import {
  toYMD,
  newState,
  parseState,
  decide,
  findResponse,
  isBounce,
  renderTemplate,
} from "./logic.js";

const MSG_FIELDS =
  "id,subject,sentDateTime,conversationId,toRecipients,ccRecipients,from,isDraft,webLink,categories";

const expandProp = (id) => `$expand=singleValueExtendedProperties($filter=id eq '${q(id)}')`;
const propValue = (msg, id) =>
  (msg.singleValueExtendedProperties || []).find((p) => p.id.toLowerCase() === id.toLowerCase())?.value;

// ---------------------------------------------------------------- identity
let meCache = null;
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

const isMine = (me, msg) => me.addresses.has((msg.from?.emailAddress?.address || "").toLowerCase());

// ---------------------------------------------------------------- state writes
async function writeState(msg, state, { tracked }) {
  const cats = new Set(msg.categories || []);
  if (tracked) cats.add(CATEGORY);
  else cats.delete(CATEGORY);
  await graph("PATCH", `/me/messages/${msg.id}`, {
    categories: [...cats],
    singleValueExtendedProperties: [{ id: STATE_PROP_ID, value: JSON.stringify(state) }],
  });
}

export async function startTracking(msg, days, holidays) {
  const state = newState(days, msg.sentDateTime, holidays);
  await writeState(msg, state, { tracked: true });
  return state;
}

export async function stopTracking(msgId, status = "cancelled") {
  const msg = await graph("GET", `/me/messages/${msgId}?$select=id,categories&${expandProp(STATE_PROP_ID)}`);
  const state = parseState(propValue(msg, STATE_PROP_ID)) || { v: 1 };
  await writeState(msg, { ...state, status }, { tracked: false });
}

// ---------------------------------------------------------------- read-mode "Track this"
export async function findSentCopy(internetMessageId) {
  const me = await getMe();
  const hits = await graphAll(
    `/me/messages?$filter=internetMessageId eq '${q(internetMessageId)}'&$select=${MSG_FIELDS}&${expandProp(
      STATE_PROP_ID
    )}`,
    10
  );
  return hits.find((m) => !m.isDraft && isMine(me, m)) || null;
}

export async function trackSentEmail(internetMessageId, days, holidays) {
  const msg = await findSentCopy(internetMessageId);
  if (!msg) throw new Error("Only emails you have sent can be tracked. Open it from Sent Items and try again.");
  return startTracking(msg, days, holidays);
}

// ---------------------------------------------------------------- discovery of compose-marked emails
async function discoverComposeMarked(settings, holidays) {
  const since = settings.scanFrom || new Date(Date.now() - settings.scanDays * 864e5).toISOString();
  const path =
    `/me/mailFolders/sentitems/messages?$filter=sentDateTime ge ${since}` +
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
    await startTracking(m, days, holidays);
    // Mark as picked up so a later scan doesn't restart tracking after a reply.
    const { fuDays, ...rest } = props;
    await graph("PATCH", `/me/messages/${m.id}`, {
      singleValueExtendedProperties: [{ id: CECP_PROP_ID, value: JSON.stringify({ ...rest, fuPromoted: "1" }) }],
    });
    found++;
  }
  // Overlap by 2 days so late-arriving Sent Items aren't missed. Roams to the other Mac.
  const scanFrom = new Date(Date.now() - 2 * 864e5).toISOString();
  return { found, scanFrom };
}

// ---------------------------------------------------------------- follow-up draft
function escapeHtml(s) {
  return s.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]);
}

async function createFollowUpDraft(original, number, template, me) {
  const notMe = (r) => !me.addresses.has((r.emailAddress?.address || "").toLowerCase());
  let to = (original.toRecipients || []).filter(notMe);
  let cc = (original.ccRecipients || []).filter(notMe);
  if (!to.length && cc.length) [to, cc] = [cc, []];

  // Forward (not reply) so the original attachments come along. Stays in the same conversation.
  const draft = await graph("POST", `/me/messages/${original.id}/createForward`, {
    message: { toRecipients: to, ccRecipients: cc },
  });

  const full = await graph("GET", `/me/messages/${draft.id}?$select=body`);
  const intro =
    `<div style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:11pt">` +
    escapeHtml(renderTemplate(template, number)).replace(/\n/g, "<br>") +
    `</div><br>`;
  let html = full.body?.content || "";
  html = /<body[^>]*>/i.test(html) ? html.replace(/<body[^>]*>/i, (m) => m + intro) : intro + html;

  // Clear categories so the sent follow-up itself is never treated as a new tracked email.
  await graph("PATCH", `/me/messages/${draft.id}`, {
    body: { contentType: "html", content: html },
    categories: [],
  });
  return draft;
}

// ---------------------------------------------------------------- the check
/**
 * Runs on demand only ("Check follow-ups now"). For each tracked email:
 *  - not yet due            → no mailbox query at all
 *  - due, response found    → tracking stops, category removed
 *  - due, no response       → Follow Up #N draft created in the same conversation
 *  - draft already waiting  → reminds you to send it
 */
export async function checkAll(settings, holidays, onProgress = () => {}) {
  const me = await getMe();
  const today = toYMD(new Date());

  onProgress("Looking for newly sent emails marked for follow-up…");
  const discovery = await discoverComposeMarked(settings, holidays);

  onProgress("Loading tracked emails…");
  const tracked = await graphAll(
    `/me/messages?$filter=categories/any(c:c eq '${q(CATEGORY)}')&$select=${MSG_FIELDS}&${expandProp(
      STATE_PROP_ID
    )}&$top=100`,
    500
  );

  const results = [];
  let i = 0;
  for (const msg of tracked) {
    i++;
    if (msg.isDraft) continue;
    onProgress(`Checking ${i} of ${tracked.length}…`);
    const base = { id: msg.id, subject: msg.subject, to: msg.toRecipients, webLink: msg.webLink };
    try {
      if (!isMine(me, msg)) {
        results.push({ ...base, kind: "ignored", note: "Not sent by you – remove the Follow-up category." });
        continue;
      }
      let state = parseState(propValue(msg, STATE_PROP_ID));
      // Category added by hand in Outlook, or re-added after a reply/cancel → (re)start.
      if (!state || state.status !== "active") {
        const startFrom = state ? new Date().toISOString() : msg.sentDateTime;
        state = await startTracking({ ...msg, sentDateTime: startFrom }, state?.days || settings.defaultDays, holidays);
      }

      const ctx = {};
      if (state.draftId) {
        const d = await graph("GET", `/me/messages/${state.draftId}?$select=id,isDraft,sentDateTime,webLink`, undefined, {
          allow404: true,
        });
        ctx.draft = d ? { exists: true, isDraft: d.isDraft, sentDateTime: d.sentDateTime, webLink: d.webLink } : { exists: false };
      }

      const due = state.nextDue <= today || !!state.draftId;
      if (!due) {
        results.push({ ...base, kind: "waiting", state });
        continue;
      }

      const conv = await graphAll(
        `/me/messages?$filter=conversationId eq '${q(msg.conversationId)}'` +
          `&$select=id,from,subject,receivedDateTime,sentDateTime,isDraft&$top=50`,
        300
      );
      const response = findResponse(conv, msg.sentDateTime, me.addresses);
      ctx.responded = !!response;
      const bounced = conv.some((m) => isBounce(m) && new Date(m.receivedDateTime) > new Date(msg.sentDateTime));

      const r = decide(state, ctx, today, holidays);
      if (r.action === "stop") {
        await writeState(msg, r.state, { tracked: false });
        results.push({ ...base, kind: "replied", by: response?.from?.emailAddress?.name || response?.from?.emailAddress?.address });
      } else if (r.action === "waitDraft") {
        results.push({ ...base, kind: "draftWaiting", state, draftId: state.draftId, draftLink: ctx.draft?.webLink, bounced });
      } else if (r.action === "createDraft") {
        const draft = await createFollowUpDraft(msg, r.number, settings.template, me);
        const next = { ...r.state, draftId: draft.id };
        await writeState(msg, next, { tracked: true });
        results.push({ ...base, kind: "draftCreated", number: r.number, state: next, draftId: draft.id, draftLink: draft.webLink, bounced });
      } else {
        await writeState(msg, r.state, { tracked: true });
        results.push({ ...base, kind: "waiting", state: r.state });
      }
    } catch (e) {
      results.push({ ...base, kind: "error", note: e.message });
    }
  }
  return { results, discovered: discovery.found, scanFrom: discovery.scanFrom };
}

/** Plain REST id (non-immutable) for opening an item with Office.js. */
export async function restIdFor(immutableId) {
  const m = await graph("GET", `/me/messages/${immutableId}?$select=id,webLink`, undefined, { immutable: false });
  return m;
}
