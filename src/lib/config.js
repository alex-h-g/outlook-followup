// ---------------------------------------------------------------------------
// One entry per Microsoft 365 organisation (tenant) the add-in is used in.
// The add-in picks the entry whose domain matches the signed-in Outlook account.
// Each organisation needs its own app registration (README §2).
// Neither value is a secret – they identify the app, they do not grant access.
// ---------------------------------------------------------------------------
export const TENANTS = [
  {
    name: "GSK Electronics",
    domains: ["gsk.com.sg"],
    clientId: "3ac91786-9334-41aa-a056-a741343b95a4",
    tenantId: "26e19ffc-5943-419e-b59f-d00d4fda6259",
  },
  {
    name: "SILQ Law",
    domains: ["silqlaw.com"],
    clientId: "9ca15ed4-1e21-40aa-90c5-117368a80de9",
    tenantId: "322c38bc-a7e3-4336-a614-574780ae56a4",
  },
  {
    name: "FOSR Law",
    domains: ["fosrlaw.com"],
    clientId: "49d0a3a9-3e19-43d4-a24a-6405960aac35",
    tenantId: "af62a6e3-2d22-4a83-89f4-a08ab66aca74",
  },
  {
    name: "Hayden-Gilbert",
    domains: ["hayden-gilbert.com"],
    clientId: "2f5999df-b36b-479d-b6ac-39214f817dca",
    tenantId: "ebe589f4-4f52-400e-af00-4bbb77fd21d5",
  },
];

// ---------------------------------------------------------------------------
// Do not change below unless you know why. Changing ADDIN_ID orphans emails
// that were marked in compose but not yet picked up by "Check follow-ups".
// ---------------------------------------------------------------------------

// Must equal <Id> in manifest.xml. Office.js stores compose-time custom
// properties in a MAPI named property called "cecp-<ADDIN_ID>".
export const ADDIN_ID = "1eecab5a-ef46-49fc-98de-b06c0b6e3ca6";

// Namespace GUID for our own tracking-state property on sent emails.
const STATE_NS = "3297e62a-f98e-47e2-8b3f-3e5723f88db7";

export const CECP_PROP_ID = `String {00020329-0000-0000-C000-000000000046} Name cecp-${ADDIN_ID}`;
export const STATE_PROP_ID = `String {${STATE_NS}} Name FollowUpState`;

// Visible Outlook category on tracked sent emails. Removing it in Outlook = stop tracking.
export const CATEGORY = "Follow-up";

// Least-privilege delegated Graph scopes. No Mail.Send: you always send drafts yourself.
// Mail.ReadWrite.Shared = shared mailboxes you already have Full Access to (nothing more).
export const SCOPES = ["User.Read", "Mail.ReadWrite", "Mail.ReadWrite.Shared"];

export const DEFAULT_SETTINGS = {
  template:
    "Follow Up #{n}\n\nI'm just following up on my email below. I'd be grateful for an update when you have a moment.",
  holidays: [], // ["2026-12-31", ...] – YYYY-MM-DD, local dates
  defaultDays: 3, // used when you add the "Follow-up" category manually in Outlook
  scanDays: 30, // how far back the first scan of Sent Items looks
  mailboxes: ["me"], // "me" + shared mailbox addresses; added automatically when you use the add-in there
  scanFrom: {}, // per mailbox
  pending: [], // follow-ups chosen in a shared mailbox, waiting to be matched in Sent Items
};
