import { test } from "node:test";
import assert from "node:assert/strict";
import {
  addBusinessDays,
  toYMD,
  findResponse,
  isAutoMessage,
  newState,
  decide,
  renderTemplate,
  parseHolidays,
} from "../src/lib/logic.js";

const d = (s) => new Date(s + "T10:00:00");

test("business days skip weekends", () => {
  assert.equal(addBusinessDays(d("2026-11-13"), 3), "2026-11-18"); // Fri → Wed
  assert.equal(addBusinessDays(d("2026-11-16"), 5), "2026-11-23"); // Mon → next Mon
  assert.equal(addBusinessDays(d("2026-11-14"), 3), "2026-11-18"); // Sat → Wed
});

test("business days skip holidays", () => {
  assert.equal(addBusinessDays(d("2026-12-29"), 3, ["2026-12-31", "2027-01-01"]), "2027-01-05");
});

const me = new Set(["alex@fosrlaw.com"]);
const msg = (from, when, extra = {}) => ({
  from: { emailAddress: { address: from } },
  receivedDateTime: when,
  isDraft: false,
  subject: "RE: Contract",
  ...extra,
});

test("response detection", () => {
  const sent = "2026-11-10T03:00:00Z";
  assert.equal(findResponse([msg("alex@fosrlaw.com", "2026-11-11T00:00:00Z")], sent, me), null, "own messages");
  assert.equal(findResponse([msg("bob@x.com", "2026-11-09T00:00:00Z")], sent, me), null, "earlier messages");
  assert.equal(
    findResponse([msg("bob@x.com", "2026-11-11T00:00:00Z", { subject: "Automatic reply: Contract" })], sent, me),
    null,
    "auto-reply"
  );
  assert.equal(findResponse([msg("bob@x.com", "2026-11-11T00:00:00Z", { isDraft: true })], sent, me), null, "draft");
  assert.ok(findResponse([msg("BOB@x.com", "2026-11-11T00:00:00Z")], sent, me), "real reply");
});

test("auto message patterns", () => {
  for (const s of ["Automatic reply: x", "Out of Office: x", "Undeliverable: x", "Read: x"])
    assert.ok(isAutoMessage({ subject: s, from: { emailAddress: { address: "a@b.c" } } }), s);
  assert.ok(isAutoMessage({ subject: "x", from: { emailAddress: { address: "MAILER-DAEMON@b.c" } } }));
  assert.ok(!isAutoMessage({ subject: "RE: Out of office plans", from: { emailAddress: { address: "a@b.c" } } }));
});

test("cycle: due → draft → sent → next cycle → reply stops", () => {
  let s = newState(3, "2026-11-13T03:00:00Z", []); // Fri
  assert.equal(s.nextDue, "2026-11-18");

  let r = decide(s, { responded: false }, "2026-11-17");
  assert.equal(r.action, "notDue");

  r = decide(s, { responded: false }, "2026-11-18");
  assert.equal(r.action, "createDraft");
  assert.equal(r.number, 1);
  s = { ...r.state, draftId: "D1" };

  r = decide(s, { responded: false, draft: { exists: true, isDraft: true } }, "2026-11-19");
  assert.equal(r.action, "waitDraft");

  // Follow Up #1 sent Thu 19th → next due Tue 24th
  r = decide(s, { responded: false, draft: { exists: true, isDraft: false, sentDateTime: "2026-11-19T04:00:00Z" } }, "2026-11-20");
  assert.equal(r.action, "notDue");
  assert.equal(r.state.sent, 1);
  assert.equal(r.state.nextDue, "2026-11-24");
  assert.equal(r.state.draftId, null);
  s = r.state;

  r = decide(s, { responded: false }, "2026-11-24");
  assert.equal(r.action, "createDraft");
  assert.equal(r.number, 2);

  r = decide(s, { responded: true }, "2026-11-24");
  assert.equal(r.action, "stop");
  assert.equal(r.state.status, "replied");
});

test("deleted draft is recreated with same number", () => {
  const s = { ...newState(3, "2026-11-13T03:00:00Z", []), draftId: "D1" };
  const r = decide(s, { responded: false, draft: { exists: false } }, "2026-11-19");
  assert.equal(r.action, "createDraft");
  assert.equal(r.number, 1);
});

test("template + holiday parsing", () => {
  assert.equal(renderTemplate("Follow Up #{n} – ping", 3), "Follow Up #3 – ping");
  assert.deepEqual(parseHolidays("2026-12-31\n bad \n2027-01-01, 2027-02-02"), ["2026-12-31", "2027-01-01", "2027-02-02"]);
  assert.equal(toYMD(new Date(2026, 0, 5)), "2026-01-05");
});

import { matchPending, normSubject, prunePending } from "../src/lib/logic.js";

test("pending marks match the right sent email", () => {
  const sent = [
    { id: "a", sentDateTime: "2026-11-17T04:00:00Z", subject: "Quote request", toRecipients: [{ emailAddress: { address: "x@y.com" } }], conversationId: "c1" },
    { id: "b", sentDateTime: "2026-11-17T04:05:00Z", subject: "RE: Permit docs", toRecipients: [{ emailAddress: { address: "p@q.com" } }], conversationId: "c2" },
    { id: "c", sentDateTime: "2026-11-17T05:00:00Z", subject: "Quote request", toRecipients: [{ emailAddress: { address: "z@y.com" } }], conversationId: "c3" },
  ];
  const pending = [
    { k: "1", at: "2026-11-17T03:58:00Z", subject: "Quote request", to: ["x@y.com"] },
    { k: "2", at: "2026-11-17T04:01:00Z", subject: "Permit docs", to: [], conv: "c2" },
    { k: "3", at: "2026-11-17T04:30:00Z", subject: "", to: ["z@y.com"] },
    { k: "4", at: "2026-11-17T06:00:00Z", subject: "Quote request", to: [] }, // nothing sent after
  ];
  const m = matchPending(pending, sent);
  assert.equal(m.get("1").id, "a");
  assert.equal(m.get("2").id, "b");
  assert.equal(m.get("3").id, "c");
  assert.equal(m.has("4"), false);
  assert.equal(matchPending(pending, sent, new Set(["a"])).get("1")?.id, "c", "already-tracked mails are skipped");
  assert.equal(normSubject("RE: FW: Hello "), "hello");
  assert.equal(prunePending([{ at: "2026-01-01T00:00:00Z" }, { at: new Date().toISOString() }]).length, 1);
});
