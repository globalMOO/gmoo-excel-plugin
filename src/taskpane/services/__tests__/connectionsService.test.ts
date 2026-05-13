import {
  loadConnections,
  saveConnections,
  createConnection,
  updateConnection,
  deleteConnection,
  findByUrlAndLabel,
  touchLastUsed,
  migrateLegacyKeyIfPresent,
  normalizeApiUrl,
} from "../connectionsService";

// jest.config.js sets OfficeRuntime: {} so OfficeRuntime.storage is undefined,
// which makes connectionsService fall through to localStorage. jsdom provides
// localStorage. Reset it between tests.
beforeEach(() => {
  localStorage.clear();
});

describe("normalizeApiUrl", () => {
  it("strips trailing slashes", () => {
    expect(normalizeApiUrl("https://app.globalmoo.com/")).toBe("https://app.globalmoo.com");
    expect(normalizeApiUrl("https://app.globalmoo.com///")).toBe("https://app.globalmoo.com");
  });

  it("trims whitespace", () => {
    expect(normalizeApiUrl("  https://x.com  ")).toBe("https://x.com");
  });
});

describe("loadConnections", () => {
  it("returns [] when storage is empty", async () => {
    expect(await loadConnections()).toEqual([]);
  });

  it("returns [] on malformed JSON", async () => {
    localStorage.setItem("vsme.connections.v1", "not json");
    expect(await loadConnections()).toEqual([]);
  });

  it("filters out entries missing required fields", async () => {
    localStorage.setItem(
      "vsme.connections.v1",
      JSON.stringify([
        { id: "a", label: "ok", apiUrl: "https://x", apiKey: "k", source: "manual", createdAt: "" },
        { id: "b", label: "missing url" }, // invalid
      ])
    );
    const list = await loadConnections();
    expect(list).toHaveLength(1);
    expect(list[0].id).toBe("a");
  });
});

describe("createConnection", () => {
  it("assigns a uuid, normalizes the URL, defaults source to manual", async () => {
    const c = await createConnection({
      label: "Cloud",
      apiUrl: "https://app.globalmoo.com/",
      apiKey: "secret",
    });
    expect(c.id).toMatch(/^[0-9a-f-]{36}$/);
    expect(c.apiUrl).toBe("https://app.globalmoo.com");
    expect(c.source).toBe("manual");
    expect(c.createdAt).toBeTruthy();
  });

  it("appends ' (2)' on label collision", async () => {
    await createConnection({ label: "Cloud", apiUrl: "https://a", apiKey: "" });
    const second = await createConnection({ label: "Cloud", apiUrl: "https://b", apiKey: "" });
    expect(second.label).toBe("Cloud (2)");
    const third = await createConnection({ label: "Cloud", apiUrl: "https://c", apiKey: "" });
    expect(third.label).toBe("Cloud (3)");
  });

  it("persists to storage", async () => {
    await createConnection({ label: "X", apiUrl: "https://x", apiKey: "k" });
    const reloaded = await loadConnections();
    expect(reloaded).toHaveLength(1);
    expect(reloaded[0].label).toBe("X");
  });
});

describe("updateConnection", () => {
  it("patches fields, preserves id and createdAt", async () => {
    const c = await createConnection({ label: "X", apiUrl: "https://x", apiKey: "k" });
    const updated = await updateConnection(c.id, { apiKey: "new" });
    expect(updated.id).toBe(c.id);
    expect(updated.createdAt).toBe(c.createdAt);
    expect(updated.apiKey).toBe("new");
  });

  it("normalizes apiUrl on update", async () => {
    const c = await createConnection({ label: "X", apiUrl: "https://x", apiKey: "k" });
    const updated = await updateConnection(c.id, { apiUrl: "https://y/" });
    expect(updated.apiUrl).toBe("https://y");
  });

  it("enforces label uniqueness on rename", async () => {
    const a = await createConnection({ label: "A", apiUrl: "https://a", apiKey: "" });
    await createConnection({ label: "B", apiUrl: "https://b", apiKey: "" });
    const renamed = await updateConnection(a.id, { label: "B" });
    expect(renamed.label).toBe("B (2)");
  });

  it("allows keeping the same label (no spurious suffix)", async () => {
    const c = await createConnection({ label: "Same", apiUrl: "https://x", apiKey: "" });
    const updated = await updateConnection(c.id, { label: "Same", apiKey: "rotated" });
    expect(updated.label).toBe("Same");
  });

  it("throws on unknown id", async () => {
    await expect(updateConnection("does-not-exist", { apiKey: "x" })).rejects.toThrow();
  });
});

describe("deleteConnection", () => {
  it("removes the matching connection", async () => {
    const a = await createConnection({ label: "A", apiUrl: "https://a", apiKey: "" });
    const b = await createConnection({ label: "B", apiUrl: "https://b", apiKey: "" });
    await deleteConnection(a.id);
    const list = await loadConnections();
    expect(list).toHaveLength(1);
    expect(list[0].id).toBe(b.id);
  });

  it("is a no-op for an unknown id", async () => {
    await createConnection({ label: "A", apiUrl: "https://a", apiKey: "" });
    await deleteConnection("does-not-exist");
    const list = await loadConnections();
    expect(list).toHaveLength(1);
  });
});

describe("findByUrlAndLabel", () => {
  it("matches normalized URL + exact label", async () => {
    await createConnection({ label: "Cloud", apiUrl: "https://app.globalmoo.com", apiKey: "k" });
    const found = await findByUrlAndLabel("https://app.globalmoo.com/", "Cloud");
    expect(found).not.toBeNull();
    expect(found!.label).toBe("Cloud");
  });

  it("returns null when nothing matches", async () => {
    await createConnection({ label: "Cloud", apiUrl: "https://app.globalmoo.com", apiKey: "k" });
    expect(await findByUrlAndLabel("https://other", "Cloud")).toBeNull();
    expect(await findByUrlAndLabel("https://app.globalmoo.com", "Other")).toBeNull();
  });
});

describe("touchLastUsed", () => {
  it("sets lastUsedAt to a fresh ISO timestamp", async () => {
    const c = await createConnection({ label: "X", apiUrl: "https://x", apiKey: "" });
    expect(c.lastUsedAt).toBeUndefined();
    await touchLastUsed(c.id);
    const [reloaded] = await loadConnections();
    expect(reloaded.lastUsedAt).toBeTruthy();
    expect(new Date(reloaded.lastUsedAt!).toString()).not.toBe("Invalid Date");
  });
});

describe("migrateLegacyKeyIfPresent", () => {
  it("creates 'Default (migrated)' from legacy keys and clears them", async () => {
    localStorage.setItem("vsme_api_key", "legacy-key");
    localStorage.setItem("vsme_api_url", "https://app.globalmoo.com/api/");
    const migrated = await migrateLegacyKeyIfPresent();
    expect(migrated).not.toBeNull();
    expect(migrated!.label).toBe("Default (migrated)");
    expect(migrated!.apiKey).toBe("legacy-key");
    expect(migrated!.apiUrl).toBe("https://app.globalmoo.com"); // /api/ suffix stripped
    expect(localStorage.getItem("vsme_api_key")).toBeNull();
    expect(localStorage.getItem("vsme_api_url")).toBeNull();
  });

  it("is a no-op when connections already exist", async () => {
    await createConnection({ label: "Existing", apiUrl: "https://x", apiKey: "" });
    localStorage.setItem("vsme_api_key", "should-not-migrate");
    localStorage.setItem("vsme_api_url", "https://x");
    const result = await migrateLegacyKeyIfPresent();
    expect(result).toBeNull();
    // Defensive cleanup still happens.
    expect(localStorage.getItem("vsme_api_key")).toBeNull();
    expect(localStorage.getItem("vsme_api_url")).toBeNull();
    const list = await loadConnections();
    expect(list).toHaveLength(1);
    expect(list[0].label).toBe("Existing");
  });

  it("returns null when nothing to migrate", async () => {
    expect(await migrateLegacyKeyIfPresent()).toBeNull();
  });

  it("is idempotent on second call", async () => {
    localStorage.setItem("vsme_api_key", "legacy");
    localStorage.setItem("vsme_api_url", "https://x");
    await migrateLegacyKeyIfPresent();
    const second = await migrateLegacyKeyIfPresent();
    expect(second).toBeNull();
    const list = await loadConnections();
    expect(list).toHaveLength(1);
  });

  it("handles missing key (URL only)", async () => {
    localStorage.setItem("vsme_api_url", "https://x");
    const migrated = await migrateLegacyKeyIfPresent();
    expect(migrated).not.toBeNull();
    expect(migrated!.apiKey).toBe("");
    expect(migrated!.apiUrl).toBe("https://x");
  });
});

describe("saveConnections round-trip", () => {
  it("preserves all fields", async () => {
    const a = await createConnection({
      label: "Round Trip",
      apiUrl: "https://x",
      apiKey: "secret",
      source: "activation",
    });
    const list = await loadConnections();
    expect(list[0]).toEqual(a);
  });

  it("saveConnections([]) wipes the list", async () => {
    await createConnection({ label: "A", apiUrl: "https://a", apiKey: "" });
    await saveConnections([]);
    expect(await loadConnections()).toEqual([]);
  });
});
