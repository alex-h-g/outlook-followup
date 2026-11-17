// Microsoft sign-in via MSAL nested app authentication (NAA).
// Outlook itself brokers the sign-in, so you normally never see a login prompt
// after the first consent. Tokens stay in memory / MSAL's cache – this add-in
// never logs, stores or transmits them anywhere except to graph.microsoft.com.
import { createNestablePublicClientApplication } from "@azure/msal-browser";
import { TENANTS, SCOPES } from "./config.js";

let pcaPromise = null;

/** Tenant entry for the Outlook account this add-in is running in. */
export function currentTenant() {
  const email = (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase();
  const domain = email.split("@")[1] || "";
  const t = TENANTS.find((x) => x.domains.some((d) => domain === d || domain.endsWith("." + d)));
  if (!t) {
    throw new Error(
      `This add-in isn't set up for ${domain || "this account"} yet. Add it to TENANTS in src/lib/config.js (README §2).`
    );
  }
  return t;
}

function pca() {
  if (!pcaPromise) {
    const t = currentTenant();
    pcaPromise = createNestablePublicClientApplication({
      auth: {
        clientId: t.clientId,
        authority: `https://login.microsoftonline.com/${t.tenantId}`,
      },
      cache: { cacheLocation: "sessionStorage" },
    });
  }
  return pcaPromise;
}

export function naaSupported() {
  return !!(window.Office && Office.context.requirements.isSetSupported("NestedAppAuth", "1.1"));
}

export async function getToken() {
  if (!naaSupported()) {
    throw new Error(
      "This Outlook build doesn't support nested app authentication. Update Outlook for Mac (Help → Check for Updates)."
    );
  }
  const app = await pca();
  const request = { scopes: SCOPES };
  try {
    const r = await app.acquireTokenSilent(request);
    return r.accessToken;
  } catch {
    // First run or consent needed: Outlook shows Microsoft's consent dialog.
    const r = await app.acquireTokenPopup(request);
    return r.accessToken;
  }
}
