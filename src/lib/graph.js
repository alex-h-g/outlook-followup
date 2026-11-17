// Minimal Microsoft Graph client: delegated token, your mailbox only (/me).
import { getToken } from "./auth.js";

const BASE = "https://graph.microsoft.com/v1.0";

export class GraphError extends Error {
  constructor(status, code, message) {
    super(`Graph ${status}${code ? " " + code : ""}: ${message}`);
    this.status = status;
    this.code = code;
  }
}

/**
 * @param opts.immutable  send Prefer: IdType="ImmutableId" so message IDs survive
 *                        moves (draft → Sent Items). Default true.
 * @param opts.allow404   return null instead of throwing on 404.
 */
export async function graph(method, path, body, opts = {}) {
  const { immutable = true, allow404 = false } = opts;
  const url = path.startsWith("http") ? path : BASE + path;
  for (let attempt = 0; attempt < 4; attempt++) {
    const token = await getToken();
    const headers = { Authorization: `Bearer ${token}` };
    if (immutable) headers.Prefer = 'IdType="ImmutableId"';
    if (body !== undefined) headers["Content-Type"] = "application/json";
    const res = await fetch(url, {
      method,
      headers,
      body: body !== undefined ? JSON.stringify(body) : undefined,
    });
    if (res.status === 429 || res.status === 503) {
      const wait = Number(res.headers.get("Retry-After") || 2 ** attempt);
      await new Promise((r) => setTimeout(r, wait * 1000));
      continue;
    }
    if (res.status === 404 && allow404) return null;
    if (res.status === 204) return {};
    const text = await res.text();
    const json = text ? JSON.parse(text) : {};
    if (!res.ok) throw new GraphError(res.status, json?.error?.code, json?.error?.message || res.statusText);
    return json;
  }
  throw new GraphError(429, "TooManyRequests", "Graph kept throttling; try again in a minute.");
}

/** Follow @odata.nextLink up to `max` items. */
export async function graphAll(path, max = 500) {
  const out = [];
  let next = path;
  while (next && out.length < max) {
    const page = await graph("GET", next);
    out.push(...(page.value || []));
    next = page["@odata.nextLink"];
  }
  return out.slice(0, max);
}

export const q = (s) => encodeURIComponent(s);
