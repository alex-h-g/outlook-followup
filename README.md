# Outlook Follow-Up Tracker (Outlook for Mac · Microsoft 365)

This is a small personal Outlook add-in. You mark emails you send for follow-up. Later, you press **Check follow-ups now**, on either Mac. For each tracked email that is due and hasn't had a reply, the add-in creates a **Follow Up #N** draft in the same conversation. You review the draft and send it yourself.

- **No server or database, and IT doesn't install anything.** The static files are served from GitHub Pages. All tracking state is stored on your own sent emails in your mailbox.
- **Uses Microsoft Graph for your mailbox only.** It asks for delegated `Mail.ReadWrite` and `User.Read`. It never sends email on its own, so it doesn't need `Mail.Send`.
- **Works in shared mailboxes too**, for example `env-doc@gsk.com.sg`, as long as you have Full Access. Follow-up drafts are created in that mailbox, so they go out from it.
- **Works on both Macs.** Both Macs see the same state, because the state lives in the mailbox.

See [`docs/DESIGN.md`](docs/DESIGN.md) for the platform research, design decisions and open test items.

---

## 1. One-time: GitHub repo + Pages

1. Create the repo `alex-h-g/outlook-followup` (public). Push this code to it.
2. In the repo, open **Settings → Pages → Build and deployment**, and set **Source: GitHub Actions**.
3. Each push to `main` runs the tests, builds the add-in and publishes it to
   `https://alex-h-g.github.io/outlook-followup/`.

To host somewhere else, replace `https://alex-h-g.github.io/outlook-followup` in `manifest.xml`, and the redirect URI in step 2 below.

## 2. One-time per organisation: register the app in Entra ID

The same add-in works for several Microsoft 365 organisations. Do this once for each one, then add an entry to `TENANTS` in `src/lib/config.js` with its email domain, client ID and tenant ID, and push.

| Organisation | Domain | Status |
|---|---|---|
| GSK Electronics | gsk.com.sg | Registered, admin consent granted, assignment restricted to alex@ |
| SILQ Law | – | To do |
| FOSR Law | – | To do |

If you are the tenant admin, click **Grant admin consent for <org>** on the API permissions page after step 3, and you're done.

### Steps

Most tenants let ordinary users register an app and consent for their own mailbox. If yours does, you don't need IT.

1. Go to **https://entra.microsoft.com** → **Applications → App registrations → New registration**.
   - **Name:** `Outlook Follow-Up Tracker`
   - **Supported account types:** *Accounts in this organizational directory only*
   - **Redirect URI:** platform **Single-page application (SPA)**, URI `brk-multihub://alex-h-g.github.io`
   - Click **Register**.
2. On the app's **Overview** page, copy the **Application (client) ID** and the **Directory (tenant) ID** into `src/lib/config.js`. Then commit and push.
3. Open **API permissions → Add a permission → Microsoft Graph → Delegated** and add **Mail.ReadWrite** and **Mail.ReadWrite.Shared**. **User.Read** is already listed. `.Shared` only reaches shared mailboxes you already have Full Access to, such as `env-doc@gsk.com.sg`.
4. You don't need to press "Grant admin consent". The first time you use **Check follow-ups now**, Microsoft shows a consent prompt for your own account.

**If "New registration" is blocked**, or the first run says **"Need admin approval"**, your tenant doesn't allow self-service. Send IT this:

> Please allow (or register and consent to) a single-tenant SPA app for my own use only. Delegated Microsoft Graph permissions: `User.Read` and `Mail.ReadWrite`. No application permissions, no access to other mailboxes. SPA redirect URI: `brk-multihub://alex-h-g.github.io`.

## 3. Install the add-in in Outlook (sideload, once per account)

1. In a browser, open **https://aka.ms/olksideload**. This opens Outlook on the web at *Add-ins for Outlook*.
2. Go to **My add-ins → Custom Addins → Add a custom add-in**.
   - Use **Add from URL** with `https://alex-h-g.github.io/outlook-followup/manifest.xml`.
   - If only **Add from file** is offered, download that URL first and pick the file.
3. The add-in follows your mailbox, so it appears in Outlook on both MacBooks. You may need to restart Outlook.

## 4. Using it

| Where | What |
|---|---|
| **Writing an email** → ribbon **Follow-up** | Choose *No follow-up* (default) or *1–7 business days*. |
| **Reading a sent email** → ribbon **Follow-ups** → **Track this email** | Starts tracking an email you've already sent. |
| **Any email** → **Follow-ups** → **Check follow-ups now** | Picks up newly sent marked emails. Checks the due ones for a reply, and creates **Follow Up #N** drafts in **Drafts** for those with none. |
| Outlook category **"Follow-up"** | Shows which sent emails are tracked. Add it to a sent email by hand to start tracking (uses the default interval from Settings). Remove it to stop. |
| **Settings** (inside the pane) | Follow-up text (`{n}` = number), holidays to skip, default interval, shared mailboxes to check. A shared mailbox is added automatically the first time you mark or track an email in it. Saved in your mailbox, so both Macs share them. |

**How a cycle works**
1. You send an email marked for 3 business days.
2. On or after the due date, you press **Check**. If nobody has replied, you get a *Follow Up #1* draft. This is a **forward** of the original, so the attachments are included. It's addressed to the original To/Cc, and it stays in the same conversation.
3. You send the draft. The next check counts it and schedules *Follow Up #2* for 3 business days after you sent #1.
4. Once a real reply arrives, tracking stops and the category is removed. Out-of-office replies, bounces and read receipts don't count as replies.

Emails are only looked at when they're due, and only when you press **Check**. Nothing runs in the background.

## 5. Development

```bash
cd ~/outlook-followup && npm install && npm test && npm run build
```

Local testing on your Mac (sideload `manifest.xml` after replacing the GitHub Pages URL with `https://localhost:3000`):

```bash
cd ~/outlook-followup && npx --yes office-addin-dev-certs install && npm run build && npm run serve
```

Release a change:

```bash
cd ~/outlook-followup && npm test && git add -A && git commit -m "Describe change" && git push
```

If you change anything in `manifest.xml`, increase `<Version>` and re-add the add-in (step 3).
