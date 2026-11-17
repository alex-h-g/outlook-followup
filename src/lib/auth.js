// Microsoft sign-in via MSAL nested app authentication (NAA).
// Outlook itself brokers the sign-in, so you normally never see a login prompt
// after the first consent. Tokens stay in memory / MSAL's cache – this add-in
// never logs, stores or transmits them anywhere except to graph.microsoft.com.
import { createNestablePublicClientApplication } from "@azure/msal-browser";
import { CLIENT_ID, TENANT_ID, SCOPES } from "./config.js";

let pcaPromise = null;

function pca() {
  if (!pcaPromise) {
    if (CLIENT_ID.startsWith("00000000")) {
      throw new Error("CLIENT_ID / TENANT_ID not set in src/lib/config.js – see README §2.");
    }
    pcaPromise = createNestablePublicClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: `https://login.microsoftonline.com/${TENANT_ID}`,
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
