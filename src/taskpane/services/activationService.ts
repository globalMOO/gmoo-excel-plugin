// Activation flow: exchange a one-time token from a deep link for a Connection.
// See Project-Spec.md §3 for the full contract. The same exchange endpoint is
// served by both the SaaS API server and any enterprise on-prem instance — the
// add-in does not branch on deployment type.

import { createConnection, findByUrlAndLabel, normalizeApiUrl, updateConnection } from "./connectionsService";
import type { Connection } from "../types/connection";

const EXCHANGE_PATH = "/api/v1/activation/exchange";

export interface ActivationParams {
  token: string;
  srv: string;
  label?: string;
}

export interface ActivationExchangeResponse {
  apiKey: string;
  apiUrl: string;
  suggestedLabel?: string;
  userId?: string;
  expiresAt?: string | null;
}

export class ActivationError extends Error {
  constructor(
    public code:
      | "invalid_srv"
      | "network"
      | "not_found"
      | "already_used"
      | "expired"
      | "url_mismatch"
      | "malformed_response"
      | "server_error",
    message: string,
    public httpStatus?: number
  ) {
    super(message);
    this.name = "ActivationError";
  }
}

/**
 * Read activation params from window.location.search. Returns null if no
 * activation token is present (the normal case).
 */
export function parseActivationFromUrl(): ActivationParams | null {
  if (typeof window === "undefined" || !window.location) return null;
  const params = new URLSearchParams(window.location.search);
  const token = params.get("activation");
  const srv = params.get("srv");
  if (!token || !srv) return null;
  const label = params.get("label") ?? undefined;
  return { token, srv, label };
}

/**
 * Strip the activation params from the URL so a refresh doesn't re-trigger the
 * exchange (one-time tokens won't redeem twice anyway, but a stale error
 * banner would be confusing).
 */
export function clearActivationFromUrl(): void {
  if (typeof window === "undefined" || !window.history?.replaceState) return;
  const url = new URL(window.location.href);
  url.searchParams.delete("activation");
  url.searchParams.delete("srv");
  url.searchParams.delete("label");
  window.history.replaceState(null, "", url.toString());
}

/**
 * Validate that srv is a syntactically well-formed https:// URL. http:// is
 * rejected per spec §8.3.
 */
function validateSrv(srv: string): URL {
  let parsed: URL;
  try {
    parsed = new URL(srv);
  } catch {
    throw new ActivationError("invalid_srv", `Activation URL is malformed: ${srv}`);
  }
  if (parsed.protocol !== "https:") {
    throw new ActivationError("invalid_srv", `Activation URL must use https: (got ${parsed.protocol})`);
  }
  return parsed;
}

/**
 * POST {srv}/api/v1/activation/exchange with the token. Returns the parsed
 * response or throws an ActivationError. The response's apiUrl is verified
 * against srv (defense in depth: a tampered link can't smuggle a different
 * server URL into the user's connection list).
 */
export async function exchangeActivation(
  srv: string,
  token: string
): Promise<ActivationExchangeResponse> {
  validateSrv(srv);
  const base = normalizeApiUrl(srv);
  const url = `${base}${EXCHANGE_PATH}`;

  let response: Response;
  try {
    response = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json", Accept: "application/json" },
      body: JSON.stringify({ token }),
    });
  } catch (err) {
    throw new ActivationError(
      "network",
      `Could not reach activation server at ${base}: ${err instanceof Error ? err.message : String(err)}`
    );
  }

  if (response.status === 404) {
    throw new ActivationError("not_found", "Activation token not recognized.", 404);
  }
  if (response.status === 410) {
    let reason: string | undefined;
    try {
      const body = await response.json();
      reason = body?.reason;
    } catch {
      // body not JSON
    }
    if (reason === "already_used") {
      throw new ActivationError("already_used", "This activation link has already been used.", 410);
    }
    throw new ActivationError("expired", "This activation link has expired.", 410);
  }
  if (!response.ok) {
    throw new ActivationError(
      "server_error",
      `Activation server returned ${response.status}.`,
      response.status
    );
  }

  let body: ActivationExchangeResponse;
  try {
    body = (await response.json()) as ActivationExchangeResponse;
  } catch {
    throw new ActivationError("malformed_response", "Activation response was not valid JSON.");
  }

  if (typeof body.apiKey !== "string" || typeof body.apiUrl !== "string" || !body.apiKey || !body.apiUrl) {
    throw new ActivationError("malformed_response", "Activation response is missing apiKey or apiUrl.");
  }

  // Echo-back check: server must affirm the URL the user was directed to.
  if (normalizeApiUrl(body.apiUrl) !== normalizeApiUrl(srv)) {
    throw new ActivationError(
      "url_mismatch",
      `Activation server URL mismatch: expected ${srv}, got ${body.apiUrl}`
    );
  }

  return body;
}

/**
 * Apply an exchange result to the connection list:
 *   - If a connection with the same (apiUrl, label) already exists, update its
 *     apiKey + lastUsedAt. (Re-activation case.)
 *   - Otherwise create a new Connection with source="activation".
 *
 * Label preference order (spec §3.1): explicit `label` arg → server's
 * suggestedLabel → derived from URL hostname.
 */
export async function applyActivation(
  result: ActivationExchangeResponse,
  preferredLabel?: string
): Promise<Connection> {
  const apiUrl = normalizeApiUrl(result.apiUrl);
  const label = preferredLabel?.trim() || result.suggestedLabel?.trim() || deriveLabelFromUrl(apiUrl);

  const existing = await findByUrlAndLabel(apiUrl, label);
  if (existing) {
    return updateConnection(existing.id, {
      apiKey: result.apiKey,
      lastUsedAt: new Date().toISOString(),
    });
  }
  return createConnection({
    label,
    apiUrl,
    apiKey: result.apiKey,
    source: "activation",
  });
}

function deriveLabelFromUrl(url: string): string {
  try {
    const host = new URL(url).hostname;
    if (host === "app.globalmoo.com") return "globalMOO Cloud";
    return host;
  } catch {
    return "globalMOO";
  }
}
