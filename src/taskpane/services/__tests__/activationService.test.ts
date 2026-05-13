import {
  parseActivationFromUrl,
  clearActivationFromUrl,
  exchangeActivation,
  applyActivation,
  ActivationError,
} from "../activationService";
import { loadConnections, createConnection } from "../connectionsService";

beforeEach(() => {
  localStorage.clear();
  // Reset URL between tests. jsdom defaults to about:blank; use a same-origin
  // path so subsequent replaceState calls in tests don't trip the same-origin
  // security check.
  window.history.replaceState(null, "", "/");
});

describe("parseActivationFromUrl", () => {
  it("returns null when no token is present", () => {
    expect(parseActivationFromUrl()).toBeNull();
  });

  it("returns null when token is present but srv is missing", () => {
    window.history.replaceState(null, "", "?activation=abc");
    expect(parseActivationFromUrl()).toBeNull();
  });

  it("parses token, srv, and optional label", () => {
    window.history.replaceState(
      null,
      "",
      "?activation=tok123&srv=" + encodeURIComponent("https://app.globalmoo.com") + "&label=Cloud"
    );
    const parsed = parseActivationFromUrl();
    expect(parsed).toEqual({ token: "tok123", srv: "https://app.globalmoo.com", label: "Cloud" });
  });

  it("omits label when absent", () => {
    window.history.replaceState(
      null,
      "",
      "?activation=tok&srv=" + encodeURIComponent("https://x.com")
    );
    expect(parseActivationFromUrl()?.label).toBeUndefined();
  });
});

describe("clearActivationFromUrl", () => {
  it("strips activation params, preserves the rest", () => {
    window.history.replaceState(
      null,
      "",
      "/taskpane.html?activation=tok&srv=https%3A%2F%2Fx&label=L&keep=me"
    );
    clearActivationFromUrl();
    const params = new URLSearchParams(window.location.search);
    expect(params.get("activation")).toBeNull();
    expect(params.get("srv")).toBeNull();
    expect(params.get("label")).toBeNull();
    expect(params.get("keep")).toBe("me");
  });
});

describe("exchangeActivation", () => {
  const realFetch = global.fetch;
  afterEach(() => {
    global.fetch = realFetch;
  });

  // Build a Response-like object. We avoid `new Response(...)` because the
  // jsdom version pinned by this project doesn't ship the fetch Response
  // constructor. The activationService only uses .status, .ok, .json(), .text().
  function makeResponse(status: number, body?: unknown) {
    return {
      status,
      ok: status >= 200 && status < 300,
      async json() {
        if (body === undefined) throw new Error("no body");
        return body;
      },
      async text() {
        return body === undefined ? "" : JSON.stringify(body);
      },
    };
  }

  function mockFetch(impl: (url: string, init?: RequestInit) => Promise<unknown>) {
    global.fetch = jest.fn(impl) as unknown as typeof fetch;
  }

  function jsonResponse(status: number, body: unknown) {
    return makeResponse(status, body);
  }

  it("rejects http:// URLs", async () => {
    await expect(exchangeActivation("http://insecure.com", "tok")).rejects.toMatchObject({
      code: "invalid_srv",
    });
  });

  it("rejects malformed URLs", async () => {
    await expect(exchangeActivation("not-a-url", "tok")).rejects.toMatchObject({
      code: "invalid_srv",
    });
  });

  it("POSTs to {srv}/api/v1/activation/exchange with the token", async () => {
    mockFetch(async (url, init) => {
      expect(url).toBe("https://app.globalmoo.com/api/v1/activation/exchange");
      expect(init?.method).toBe("POST");
      expect(JSON.parse(init?.body as string)).toEqual({ token: "tok" });
      return jsonResponse(200, {
        apiKey: "gm_live_abc",
        apiUrl: "https://app.globalmoo.com",
        suggestedLabel: "Cloud",
      });
    });
    const res = await exchangeActivation("https://app.globalmoo.com", "tok");
    expect(res.apiKey).toBe("gm_live_abc");
  });

  it("normalizes srv (handles trailing slash)", async () => {
    mockFetch(async (url) => {
      expect(url).toBe("https://app.globalmoo.com/api/v1/activation/exchange");
      return jsonResponse(200, { apiKey: "k", apiUrl: "https://app.globalmoo.com" });
    });
    await exchangeActivation("https://app.globalmoo.com/", "tok");
  });

  it("maps 404 to ActivationError(not_found)", async () => {
    mockFetch(async () => makeResponse(404));
    await expect(exchangeActivation("https://x.com", "tok")).rejects.toMatchObject({
      code: "not_found",
      httpStatus: 404,
    });
  });

  it("maps 410 with reason=already_used", async () => {
    mockFetch(async () => jsonResponse(410, { reason: "already_used" }));
    await expect(exchangeActivation("https://x.com", "tok")).rejects.toMatchObject({
      code: "already_used",
    });
  });

  it("maps 410 without reason to expired", async () => {
    mockFetch(async () => makeResponse(410));
    await expect(exchangeActivation("https://x.com", "tok")).rejects.toMatchObject({
      code: "expired",
    });
  });

  it("maps 500 to server_error", async () => {
    mockFetch(async () => makeResponse(500));
    await expect(exchangeActivation("https://x.com", "tok")).rejects.toMatchObject({
      code: "server_error",
    });
  });

  it("maps fetch rejection to network", async () => {
    mockFetch(async () => {
      throw new Error("ECONNREFUSED");
    });
    await expect(exchangeActivation("https://x.com", "tok")).rejects.toMatchObject({
      code: "network",
    });
  });

  it("rejects responses missing apiKey/apiUrl", async () => {
    mockFetch(async () => jsonResponse(200, { foo: "bar" }));
    await expect(exchangeActivation("https://x.com", "tok")).rejects.toMatchObject({
      code: "malformed_response",
    });
  });

  it("rejects when server's apiUrl doesn't match srv", async () => {
    mockFetch(async () =>
      jsonResponse(200, { apiKey: "k", apiUrl: "https://different.com" })
    );
    await expect(
      exchangeActivation("https://app.globalmoo.com", "tok")
    ).rejects.toMatchObject({ code: "url_mismatch" });
  });

  it("treats trailing-slash differences as matching", async () => {
    mockFetch(async () =>
      jsonResponse(200, { apiKey: "k", apiUrl: "https://app.globalmoo.com/" })
    );
    const res = await exchangeActivation("https://app.globalmoo.com", "tok");
    expect(res.apiKey).toBe("k");
  });
});

describe("applyActivation", () => {
  it("creates a new connection with source='activation' when none exists", async () => {
    const conn = await applyActivation({
      apiKey: "k",
      apiUrl: "https://app.globalmoo.com",
      suggestedLabel: "Cloud",
    });
    expect(conn.source).toBe("activation");
    expect(conn.label).toBe("Cloud");
    expect(conn.apiKey).toBe("k");
    expect(conn.apiUrl).toBe("https://app.globalmoo.com");
  });

  it("updates existing connection (apiKey + lastUsedAt) when (apiUrl, label) matches", async () => {
    const existing = await createConnection({
      label: "Cloud",
      apiUrl: "https://app.globalmoo.com",
      apiKey: "old",
      source: "manual",
    });
    const updated = await applyActivation(
      { apiKey: "new", apiUrl: "https://app.globalmoo.com", suggestedLabel: "Cloud" },
      "Cloud"
    );
    expect(updated.id).toBe(existing.id);
    expect(updated.apiKey).toBe("new");
    expect(updated.lastUsedAt).toBeTruthy();
    const list = await loadConnections();
    expect(list).toHaveLength(1);
  });

  it("prefers explicit label over suggestedLabel", async () => {
    const conn = await applyActivation(
      { apiKey: "k", apiUrl: "https://x.com", suggestedLabel: "Server-Suggested" },
      "User-Picked"
    );
    expect(conn.label).toBe("User-Picked");
  });

  it("derives label 'globalMOO Cloud' for app.globalmoo.com", async () => {
    const conn = await applyActivation({ apiKey: "k", apiUrl: "https://app.globalmoo.com" });
    expect(conn.label).toBe("globalMOO Cloud");
  });

  it("derives label from hostname for other URLs", async () => {
    const conn = await applyActivation({ apiKey: "k", apiUrl: "https://gmoo.acme.com" });
    expect(conn.label).toBe("gmoo.acme.com");
  });
});

describe("ActivationError", () => {
  it("carries code and httpStatus", () => {
    const err = new ActivationError("not_found", "nope", 404);
    expect(err.code).toBe("not_found");
    expect(err.httpStatus).toBe(404);
    expect(err).toBeInstanceOf(Error);
  });
});
