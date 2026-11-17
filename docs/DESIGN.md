# Design notes

These are the decisions agreed with Alex on 17 Nov 2026, and the platform facts checked that day. Read this before redeveloping.

## Requirements (as agreed)

- The only user is Alex, on two MacBooks running new Outlook for Mac with Microsoft 365 / Exchange Online. No IT involvement if it can be avoided.
- Tracking starts from:
  - a compose selector (No follow-up *(default)* / 3 / 4 / 5 business days), **and**
  - "Track this email" on an email that's already sent, **and**
  - adding the "Follow-up" category by hand.
- Checks run only when Alex presses **Check follow-ups now**. There's no polling or timer.
- If a due email has had no response, the add-in creates a **draft** that Alex reviews and sends. It never sends automatically.
- The draft is a **forward** of the original, not a reply, so attachments are kept. It's addressed to the original To + Cc (Reply All recipients, minus Alex) and stays in the same conversation. It's numbered Follow Up #1, #2, … and the cycle continues until a response arrives or Alex cancels.
- The follow-up text is an editable template.
- Hosting is GitHub Pages, from a public repo.

### Assumptions made where Alex didn't answer, to confirm or change

1. **Response:** any non-draft message in the same conversation, from anyone other than Alex, received after the tracked email was sent. Out-of-office replies, bounces and read receipts are ignored. A bounce is shown as a warning.
2. **Business days:** Monday to Friday, plus an optional holiday list in Settings (empty by default).
3. **Storage:** in the mailbox rather than on the Mac, because tracking must work on both MacBooks.
4. **Repo:** public `alex-h-g/outlook-followup`. Free GitHub Pages requires a public repo. The code contains no secrets.
5. **Draft number advances when the draft is *sent*.** If a draft is deleted unsent, the next check recreates the same number.
6. **Next due date** is N business days after the previous follow-up was actually sent.

## Platform facts verified (Nov 2026)

| Question | Finding | Consequence |
|---|---|---|
| Can Office.js alone detect replies? | No. Office.js only sees the open/selected item and has no mailbox search. | Microsoft Graph is required. |
| EWS / legacy Exchange tokens? | Microsoft started blocking EWS for non-Microsoft apps in Exchange Online on 1 Oct 2026, with final shutdown in April 2027. Legacy Exchange tokens are off. | Graph is the only supported route. |
| Auth on Outlook for Mac | MSAL nested app authentication (NAA) is GA for Outlook on Mac. It needs an Entra app registration with an SPA redirect `brk-multihub://<domain>`. | `src/lib/auth.js`. Requires a CLIENT_ID and TENANT_ID. |
| Manifest | Outlook on Mac does not support the unified (JSON) manifest. | Add-in-only XML manifest. |
| Startup event / background timer | Outlook has no "Outlook opened" event. Event handlers (OnMessageSend, OnNewMessageCompose…) must finish quickly. Task panes run only while open. | Checks run on demand (manual button). This matches Alex's choice. |
| Categories in compose | Not supported in Outlook on the web or new Outlook for Windows. Risky on new Mac, which shares that code. | Compose uses **custom properties** instead. These are documented to persist to the Sent Items copy. |
| Hosting | Add-ins need an HTTPS URL. Event-based add-ins also need internet. | GitHub Pages, static files only. |

## Architecture

```
compose.html/js     → Office.js custom property  fuDays=3|4|5  (on the draft → Sent Items copy)
taskpane.html/js    → "Track this", "Check follow-ups now", Settings
lib/tracker.js      → Graph logic (discover, list, check, create draft, stop)
lib/logic.js        → pure functions: business days, response detection, state machine (unit-tested)
lib/auth.js         → MSAL NAA token (in memory / session cache only)
lib/graph.js        → fetch wrapper, ImmutableId, 429 retry, paging
lib/office.js       → Office.js promise wrappers, roaming settings
```

### Data model (all inside Alex's mailbox)

- **Category `Follow-up`** on the sent email marks it as actively tracked. Removing it in Outlook stops tracking.
- **Extended property** `String {3297e62a-f98e-47e2-8b3f-3e5723f88db7} Name FollowUpState` holds JSON:
  `{v:1, days, sent, nextDue:"YYYY-MM-DD", draftId, status:"active|replied|cancelled", started}`
- **Office.js custom properties** `String {00020329-0000-0000-C000-000000000046} Name cecp-<AddInId>` hold `{"fuDays":"3"}` and become `{"fuPromoted":"1"}` once picked up.
- **Roaming settings** (per add-in, stored in the mailbox): `{template, holidays[], defaultDays, scanFrom}`.
- All Graph calls send `Prefer: IdType="ImmutableId"` so a draft's ID survives the move to Sent Items.

### Check algorithm

1. Scan Sent Items since `scanFrom` (first run: last 30 days) for the `cecp` marker. Start tracking any found, then mark them as picked up.
2. List messages with the `Follow-up` category.
3. For each tracked message:
   - If it isn't due and there's no pending draft, skip it. No mailbox query is made.
   - If a pending draft exists: still a draft means remind Alex; sent means count it and reschedule; deleted means forget it.
   - If it's due, fetch the conversation (`conversationId`). A response stops tracking and removes the category. Otherwise, `createForward` to the original To/Cc, prepend the template to the body, and clear the categories on the draft.
4. Save `scanFrom` = now − 2 days. The overlap is safe because each pickup is marked.

## Open items to test on first real use

1. **Compose marker survives send on new Outlook for Mac.** Mark a test email, send it, press Check, and expect "1 newly picked up". If this fails, use "Track this email" on the sent item as a workaround. The fix would be to switch the compose marker to an `OnMessageSend` event handler.
2. **Forward draft threads correctly.** The Follow Up draft should appear inside the original conversation in Outlook. Graph's `createForward` keeps the conversation, and the subject stays `FW: <original>`. The number is in the body, not the subject, because changing the subject can split the thread.
3. **User consent is allowed in the tenant.** See README §2.
4. **"Open draft" button.** It uses `displayMessageForm`. If new Outlook for Mac won't open drafts that way, it falls back to Outlook on the web. The draft is always in Drafts either way.
5. **Outlook signature isn't added to Graph-created drafts.** Add your sign-off to the template if you want one.

## Not detected (by design or by limitation)

- A reply sent as a brand-new email (a new thread) rather than a reply in the conversation.
- A reply received only by a colleague and not by you.
