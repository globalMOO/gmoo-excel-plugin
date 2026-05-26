// Pure diagnostic helpers for the Test-Connection flow. Lifted out of
// ConnectionSetup.tsx so the classification logic — which is the part the
// customer sees when something is wrong — can be unit tested without
// rendering React or mocking Fluent UI.

import { GmooApiError } from "./gmooApi";

export type FailureKind = "api" | "network";

export interface DiagnoseSuccess {
  ok: true;
}

export interface DiagnoseFailure {
  ok: false;
  failureKind: FailureKind;
  error: string;
  /**
   * Result of the follow-up no-cors reachability probe when failureKind is
   * "network". true = round-trip completed (so the original failure was CORS).
   * false = probe also failed (cert / DNS / network). null = probe skipped
   * (api failure, or caller passed probeHost = null for dev mode).
   */
  corsReachable: boolean | null;
}

export type DiagnoseResult = DiagnoseSuccess | DiagnoseFailure;

export interface DiagnoseInput {
  /** Minimal surface the diagnoser needs from GmooClient — just `getModels`. */
  client: { getModels: () => Promise<unknown> };
  /**
   * Origin to probe with a no-cors GET when the primary call throws a
   * non-GmooApiError. Pass null to skip the probe (e.g. dev mode through a
   * webpack proxy, where reachability and CORS are both masked).
   */
  probeHost: string | null;
  /** Injectable for tests. Defaults to global fetch. */
  fetchImpl?: typeof fetch;
  /** Probe budget in ms. Defaults to 5000. */
  probeTimeoutMs?: number;
}

const DEFAULT_PROBE_TIMEOUT_MS = 5000;

/**
 * Try the primary API call (`client.getModels()`). On failure, classify:
 *   - 401  → "api", "Invalid API key."
 *   - 4xx/5xx → "api", "API error (status): message"
 *   - network (TypeError etc.) → "network"; then probe `probeHost + "/api/"`
 *     with mode: no-cors to disambiguate CORS-blocked-but-reachable from
 *     truly unreachable. The probe is bounded by `probeTimeoutMs`.
 */
export async function diagnoseConnection(input: DiagnoseInput): Promise<DiagnoseResult> {
  try {
    await input.client.getModels();
    return { ok: true };
  } catch (err) {
    if (err instanceof GmooApiError && err.status === 401) {
      return { ok: false, failureKind: "api", error: "Invalid API key.", corsReachable: null };
    }
    if (err instanceof GmooApiError) {
      return {
        ok: false,
        failureKind: "api",
        error: `API error (${err.status}): ${err.apiError?.message ?? "Unknown error"}`,
        corsReachable: null,
      };
    }

    const errorMsg = err instanceof Error ? err.message : "Unknown error";
    const corsReachable = input.probeHost ? await probeReachability(input) : null;

    return { ok: false, failureKind: "network", error: errorMsg, corsReachable };
  }
}

async function probeReachability(input: DiagnoseInput): Promise<boolean> {
  const fetchImpl = input.fetchImpl ?? fetch;
  const timeout = input.probeTimeoutMs ?? DEFAULT_PROBE_TIMEOUT_MS;
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeout);
  try {
    await fetchImpl(input.probeHost + "/api/", {
      mode: "no-cors",
      cache: "no-store",
      signal: controller.signal,
    });
    return true;
  } catch {
    return false;
  } finally {
    clearTimeout(timer);
  }
}

/**
 * Build the PowerShell one-liner that hands cert-trust off to the installer.
 * The URL is single-quoted; any embedded single quote is doubled (PS escape
 * convention) so it can't terminate the string early.
 */
export function buildCertTrustCommand(apiUrl: string): string {
  const safeUrl = apiUrl.replace(/'/g, "''");
  return `& ([scriptblock]::Create((irm https://globalmoo.github.io/gmoo-excel-plugin/install.ps1))) -ApiUrl '${safeUrl}' -CertOnly`;
}
