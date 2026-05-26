import {
  diagnoseConnection,
  buildCertTrustCommand,
} from "../connectionDiagnostics";
import { GmooApiError } from "../gmooApi";

// Minimal client stub — diagnoseConnection only calls getModels().
function makeClient(behavior: () => Promise<unknown>) {
  return { getModels: behavior };
}

// Construct a fetch impl that records calls and runs caller-supplied behavior.
// diagnoseConnection doesn't inspect the Response body, so we can resolve with
// undefined instead of constructing a real Response (which is flaky in jsdom).
type FetchHandler = (input: string, init?: RequestInit) => Promise<unknown>;
function makeFetch(handler: FetchHandler): jest.Mock & typeof fetch {
  return jest.fn(handler) as unknown as jest.Mock & typeof fetch;
}

describe("diagnoseConnection — success path", () => {
  it("returns { ok: true } when getModels resolves", async () => {
    const client = makeClient(async () => []);
    const result = await diagnoseConnection({ client, probeHost: "https://x" });
    expect(result).toEqual({ ok: true });
  });
});

describe("diagnoseConnection — API failures", () => {
  it("classifies 401 as 'Invalid API key.'", async () => {
    const client = makeClient(async () => {
      throw new GmooApiError(401, null);
    });
    const result = await diagnoseConnection({ client, probeHost: "https://x" });
    expect(result).toEqual({
      ok: false,
      failureKind: "api",
      error: "Invalid API key.",
      corsReachable: null,
    });
  });

  it("formats other GmooApiError statuses with their message", async () => {
    const client = makeClient(async () => {
      throw new GmooApiError(500, {
        message: "Database is on fire",
        code: "db_fire",
      } as never);
    });
    const result = await diagnoseConnection({ client, probeHost: "https://x" });
    expect(result).toEqual({
      ok: false,
      failureKind: "api",
      error: "API error (500): Database is on fire",
      corsReachable: null,
    });
  });

  it("falls back to 'Unknown error' when GmooApiError has no apiError body", async () => {
    const client = makeClient(async () => {
      throw new GmooApiError(503, null);
    });
    const result = await diagnoseConnection({ client, probeHost: "https://x" });
    expect(result).toMatchObject({
      ok: false,
      failureKind: "api",
      error: "API error (503): Unknown error",
    });
  });

  it("does NOT run the no-cors probe for API failures", async () => {
    const probe = makeFetch(async () => undefined);
    const client = makeClient(async () => {
      throw new GmooApiError(401, null);
    });
    await diagnoseConnection({
      client,
      probeHost: "https://customer.local",
      fetchImpl: probe,
    });
    expect(probe).not.toHaveBeenCalled();
  });
});

describe("diagnoseConnection — network failures", () => {
  it("runs the probe and returns corsReachable=true when the probe resolves", async () => {
    const probe = makeFetch(async () => undefined);
    const client = makeClient(async () => {
      throw new TypeError("Failed to fetch");
    });
    const result = await diagnoseConnection({
      client,
      probeHost: "https://customer.local",
      fetchImpl: probe,
    });
    expect(result).toEqual({
      ok: false,
      failureKind: "network",
      error: "Failed to fetch",
      corsReachable: true,
    });
    expect(probe).toHaveBeenCalledWith(
      "https://customer.local/api/",
      expect.objectContaining({ mode: "no-cors", cache: "no-store" })
    );
  });

  it("returns corsReachable=false when the probe also throws", async () => {
    const probe = makeFetch(async () => {
      throw new TypeError("Failed to fetch");
    });
    const client = makeClient(async () => {
      throw new TypeError("Failed to fetch");
    });
    const result = await diagnoseConnection({
      client,
      probeHost: "https://customer.local",
      fetchImpl: probe,
    });
    expect(result).toMatchObject({
      ok: false,
      failureKind: "network",
      corsReachable: false,
    });
  });

  it("returns corsReachable=null when probeHost is null (dev mode)", async () => {
    const probe = makeFetch(async () => undefined);
    const client = makeClient(async () => {
      throw new TypeError("Failed to fetch");
    });
    const result = await diagnoseConnection({
      client,
      probeHost: null,
      fetchImpl: probe,
    });
    expect(result).toEqual({
      ok: false,
      failureKind: "network",
      error: "Failed to fetch",
      corsReachable: null,
    });
    expect(probe).not.toHaveBeenCalled();
  });

  it("aborts the probe after probeTimeoutMs and reports corsReachable=false", async () => {
    // Probe hangs until its signal aborts, then rejects with AbortError.
    const probe = makeFetch((_input, init) =>
      new Promise((_, reject) => {
        const signal = init?.signal as AbortSignal | undefined;
        if (signal) {
          signal.addEventListener("abort", () => reject(new DOMException("aborted", "AbortError")));
        }
      })
    );
    const client = makeClient(async () => {
      throw new TypeError("Failed to fetch");
    });
    const result = await diagnoseConnection({
      client,
      probeHost: "https://customer.local",
      fetchImpl: probe,
      probeTimeoutMs: 10,
    });
    expect(result).toMatchObject({
      ok: false,
      failureKind: "network",
      corsReachable: false,
    });
  });

  it("preserves a generic message when thrown value is not an Error", async () => {
    const client = makeClient(async () => {
      throw "stringly typed disaster";
    });
    const result = await diagnoseConnection({
      client,
      probeHost: null,
    });
    expect(result).toMatchObject({
      ok: false,
      failureKind: "network",
      error: "Unknown error",
    });
  });
});

describe("buildCertTrustCommand — PowerShell escaping", () => {
  it("wraps a clean URL in single quotes", () => {
    const cmd = buildCertTrustCommand("https://1.1.1.1");
    expect(cmd).toContain("-ApiUrl 'https://1.1.1.1'");
    expect(cmd).toContain("-CertOnly");
  });

  it("doubles an embedded single quote (PS escape)", () => {
    const cmd = buildCertTrustCommand("https://srv/a'b");
    expect(cmd).toContain("-ApiUrl 'https://srv/a''b'");
  });

  it("doubles every embedded single quote", () => {
    const cmd = buildCertTrustCommand("'a'b'");
    expect(cmd).toContain("-ApiUrl '''a''b'''");
  });

  it("handles URLs with no quotes unchanged", () => {
    const cmd = buildCertTrustCommand("https://app.globalmoo.com");
    expect(cmd).toContain("-ApiUrl 'https://app.globalmoo.com'");
  });

  it("does not double-quote standard URL characters (colon, slash, dot)", () => {
    // Sanity: the regex only touches single quotes.
    const cmd = buildCertTrustCommand("https://example.com:8443/path?x=1");
    expect(cmd).toContain("-ApiUrl 'https://example.com:8443/path?x=1'");
  });
});
