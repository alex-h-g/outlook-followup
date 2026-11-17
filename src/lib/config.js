// ---------------------------------------------------------------------------
// EDIT THESE TWO VALUES after registering the app in Entra ID (see README §2).
// Neither value is a secret – they identify the app, they do not grant access.
// ---------------------------------------------------------------------------
export const CLIENT_ID = "3ac91786-9334-41aa-a056-a741343b95a4"; // Application (client) ID
export const TENANT_ID = "26e19ffc-5943-419e-b59f-d00d4fda6259"; // Directory (tenant) ID

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
export const SCOPES = ["User.Read", "Mail.ReadWrite"];

export const DEFAULT_SETTINGS = {
  template:
    "Follow Up #{n}\n\nI'm just following up on my email below. I'd be grateful for an update when you have a moment.",
  holidays: [], // ["2026-12-31", ...] – YYYY-MM-DD, local dates
  defaultDays: 3, // used when you add the "Follow-up" category manually in Outlook
  scanDays: 30, // how far back the first scan of Sent Items looks
};
