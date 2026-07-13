import { GmooClient, GmooApiError } from "../gmooApi";

const BASE = "https://api.test.example/api/";

// Build a Response-like object. We avoid `new Response(...)` because the
// jsdom version pinned by this project doesn't ship the fetch Response
// constructor. The client only uses .ok, .status, .json(), .text().
function makeResponse(status: number, body?: unknown) {
  return {
    status,
    ok: status >= 200 && status < 300,
    async json() {
      if (body === undefined) throw new Error("no body");
      return body;
    },
    async text() {
      return typeof body === "string" ? body : body === undefined ? "" : JSON.stringify(body);
    },
  };
}

const realFetch = global.fetch;
let fetchMock: jest.Mock;

function mockFetchSequence(...responses: unknown[]) {
  let call = 0;
  fetchMock = jest.fn(async () => {
    const response = responses[Math.min(call, responses.length - 1)];
    call++;
    return response;
  });
  global.fetch = fetchMock as unknown as typeof fetch;
}

afterEach(() => {
  global.fetch = realFetch;
  jest.restoreAllMocks();
});

describe("constructor", () => {
  it("rejects an empty API key", () => {
    expect(() => new GmooClient("")).toThrow("API key cannot be empty.");
  });

  it("appends a trailing slash to the base URL when missing", async () => {
    mockFetchSequence(makeResponse(200, []));
    const client = new GmooClient("key", "https://api.test.example/api");
    await client.getModels();
    expect(fetchMock).toHaveBeenCalledWith(`${BASE}models`, expect.anything());
  });
});

describe("request mechanics", () => {
  it("sends Bearer auth and Accept headers on GET, with no body", async () => {
    mockFetchSequence(makeResponse(200, []));
    const client = new GmooClient("secret-key", BASE);
    await client.getModels();

    const [url, init] = fetchMock.mock.calls[0];
    expect(url).toBe(`${BASE}models`);
    expect(init.method).toBe("GET");
    expect(init.headers.Authorization).toBe("Bearer secret-key");
    expect(init.headers.Accept).toBe("application/json");
    expect(init.headers["Content-Type"]).toBeUndefined();
    expect(init.body).toBeUndefined();
  });

  it("sends a JSON body with Content-Type on POST", async () => {
    mockFetchSequence(makeResponse(200, { id: 1 }));
    const client = new GmooClient("key", BASE);
    await client.createModel("My Model", "desc");

    const [url, init] = fetchMock.mock.calls[0];
    expect(url).toBe(`${BASE}models`);
    expect(init.method).toBe("POST");
    expect(init.headers["Content-Type"]).toBe("application/json");
    expect(JSON.parse(init.body)).toEqual({ name: "My Model", description: "desc" });
  });

  it("returns the parsed JSON response", async () => {
    const model = { id: 42, name: "Solar Twin" };
    mockFetchSequence(makeResponse(200, model));
    const client = new GmooClient("key", BASE);
    await expect(client.getModel(42)).resolves.toEqual(model);
    expect(fetchMock).toHaveBeenCalledWith(`${BASE}models/42`, expect.anything());
  });
});

describe("error mapping", () => {
  it("throws GmooApiError carrying the status and parsed API error body", async () => {
    const apiError = { status: 404, title: "Not Found", message: "Model not found", errors: [] };
    mockFetchSequence(makeResponse(404, apiError));
    const client = new GmooClient("key", BASE);

    const err = await client.getModel(99).catch((e) => e);
    expect(err).toBeInstanceOf(GmooApiError);
    expect(err.status).toBe(404);
    expect(err.apiError).toEqual(apiError);
    expect(err.message).toBe("API error 404: Model not found");
  });

  it("handles a non-JSON error body without throwing a secondary error", async () => {
    mockFetchSequence(makeResponse(400, "<html>Bad Request</html>"));
    const client = new GmooClient("key", BASE);

    const err = await client.getModels().catch((e) => e);
    expect(err).toBeInstanceOf(GmooApiError);
    expect(err.status).toBe(400);
    expect(err.apiError).toBeNull();
    expect(err.message).toBe("API error 400: Unknown error");
  });

  it("handles an unreadable error body", async () => {
    mockFetchSequence({
      status: 502,
      ok: false,
      async json() { throw new Error("no body"); },
      async text() { throw new Error("stream error"); },
    });
    const client = new GmooClient("key", BASE);

    // 502 is retryable; instant delays keep the test fast
    jest.spyOn(globalThis, "setTimeout").mockImplementation(((cb: () => void) => {
      cb();
      return 0;
    }) as never);

    const err = await client.getModels().catch((e) => e);
    expect(err).toBeInstanceOf(GmooApiError);
    expect(err.status).toBe(502);
    expect(err.apiError).toBeNull();
  });
});

describe("retry with exponential backoff", () => {
  let delays: number[];

  beforeEach(() => {
    delays = [];
    jest.spyOn(globalThis, "setTimeout").mockImplementation(((cb: () => void, ms?: number) => {
      delays.push(ms ?? 0);
      cb();
      return 0;
    }) as never);
  });

  it("retries 500s with 4s/8s/10s-capped backoff, then throws after 4 attempts", async () => {
    mockFetchSequence(makeResponse(500, { message: "boom" }));
    const client = new GmooClient("key", BASE);

    await expect(client.getModels()).rejects.toMatchObject({ status: 500 });
    expect(fetchMock).toHaveBeenCalledTimes(4);
    expect(delays).toEqual([4000, 8000, 10000]);
  });

  it("retries 429 and returns the result once a retry succeeds", async () => {
    mockFetchSequence(makeResponse(429, { message: "slow down" }), makeResponse(200, [{ id: 1 }]));
    const client = new GmooClient("key", BASE);

    await expect(client.getModels()).resolves.toEqual([{ id: 1 }]);
    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(delays).toEqual([4000]);
  });

  it("does not retry client errors (400)", async () => {
    mockFetchSequence(makeResponse(400, { message: "bad request" }));
    const client = new GmooClient("key", BASE);

    await expect(client.getModels()).rejects.toMatchObject({ status: 400 });
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(delays).toEqual([]);
  });
});

describe("client-side validation (no request is sent)", () => {
  let client: GmooClient;

  beforeEach(() => {
    mockFetchSequence(makeResponse(200, {}));
    client = new GmooClient("key", BASE);
  });

  afterEach(() => {
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("rejects a non-positive model id", async () => {
    await expect(client.getModel(0)).rejects.toThrow("Model ID must be greater than zero.");
  });

  it("rejects createProject when minimums length mismatches input count", async () => {
    await expect(
      client.createProject(1, "Project", 3, [0, 0], [1, 1, 1], ["float", "float", "float"])
    ).rejects.toThrow("Length of minimums (2) does not match input count (3).");
  });

  it("rejects createProject with an unknown input type", async () => {
    await expect(
      client.createProject(1, "Project", 2, [0, 0], [1, 1], ["float", "decimal"])
    ).rejects.toThrow("Invalid input type: decimal.");
  });

  it("rejects createProject when min >= max for a float input", async () => {
    await expect(
      client.createProject(1, "Project", 2, [0, 5], [1, 5], ["float", "float"])
    ).rejects.toThrow("Minimum (5) must be less than maximum (5) for input 1.");
  });

  it("rejects loadOutputCases containing a non-finite value, pointing at Excel formulas", async () => {
    await expect(client.loadOutputCases(1, 2, [[1, NaN]])).rejects.toThrow(
      "Output case 1, value 2 is not a finite number (got NaN). Check your Excel formulas for empty cells or errors."
    );
  });

  it("rejects loadOutputCases when a case has the wrong length", async () => {
    await expect(client.loadOutputCases(1, 2, [[1, 2], [3]])).rejects.toThrow(
      "Output case 2 must be an array of length 2 (got 1)."
    );
  });

  it("rejects loadInverseOutput with an empty output list", async () => {
    await expect(client.loadInverseOutput(5, [])).rejects.toThrow("Output list cannot be empty.");
  });
});

describe("createProject boolean inputs", () => {
  it("allows min >= max for boolean inputs", async () => {
    mockFetchSequence(makeResponse(200, { id: 1 }));
    const client = new GmooClient("key", BASE);
    await expect(
      client.createProject(1, "Project", 2, [0, 0], [1, 0], ["float", "boolean"])
    ).resolves.toEqual({ id: 1 });
  });
});

describe("loadObjectives", () => {
  it("defaults exact-type bounds to zero arrays, matching the C# SDK", async () => {
    mockFetchSequence(makeResponse(200, { id: 1 }));
    const client = new GmooClient("key", BASE);
    await client.loadObjectives(3, [10, 20], ["exact", "exact"], [1, 2], [3, 4]);

    const [url, init] = fetchMock.mock.calls[0];
    expect(url).toBe(`${BASE}trials/3/objectives`);
    expect(JSON.parse(init.body)).toEqual({
      desiredL1Norm: 0,
      objectives: [10, 20],
      objectiveTypes: ["exact", "exact"],
      initialInput: [1, 2],
      initialOutput: [3, 4],
      minimumBounds: [0, 0],
      maximumBounds: [0, 0],
    });
  });

  it("leaves caller-provided bounds untouched", async () => {
    mockFetchSequence(makeResponse(200, { id: 1 }));
    const client = new GmooClient("key", BASE);
    await client.loadObjectives(3, [10], ["percent"], [1], [3], 0.5, [-2], [2]);

    const body = JSON.parse(fetchMock.mock.calls[0][1].body);
    expect(body.desiredL1Norm).toBe(0.5);
    expect(body.minimumBounds).toEqual([-2]);
    expect(body.maximumBounds).toEqual([2]);
  });
});
