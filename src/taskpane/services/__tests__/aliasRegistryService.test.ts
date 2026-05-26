import {
  loadAliases,
  loadAllAliases,
  getAlias,
  setAlias,
  clearAlias,
  clearAliasesForConnection,
  setPinned,
  isPinned,
  displayName,
} from "../aliasRegistryService";

beforeEach(() => {
  localStorage.clear();
});

describe("loadAllAliases", () => {
  it("returns [] when storage is empty", async () => {
    expect(await loadAllAliases()).toEqual([]);
  });

  it("returns [] on malformed JSON", async () => {
    localStorage.setItem("gmoo.aliases.v1", "{not json");
    expect(await loadAllAliases()).toEqual([]);
  });

  it("filters out entries with invalid shape", async () => {
    localStorage.setItem(
      "gmoo.aliases.v1",
      JSON.stringify([
        { connectionId: "c1", kind: "project", id: 1, label: "ok", updatedAt: "" },
        { connectionId: "c1", kind: "bogus", id: 2, label: "bad" },
        { id: 3 },
      ])
    );
    const list = await loadAllAliases();
    expect(list).toHaveLength(1);
    expect(list[0].id).toBe(1);
  });
});

describe("setAlias / getAlias", () => {
  it("creates a new alias", async () => {
    const a = await setAlias("conn-A", "trial", 42, "Baseline run");
    expect(a.label).toBe("Baseline run");
    expect(a.connectionId).toBe("conn-A");
    expect(a.kind).toBe("trial");
    expect(a.id).toBe(42);
    expect(a.updatedAt).toMatch(/^\d{4}-/);
  });

  it("updates an existing alias in place (no duplicates)", async () => {
    await setAlias("conn-A", "trial", 42, "First");
    await setAlias("conn-A", "trial", 42, "Second");
    const all = await loadAllAliases();
    expect(all).toHaveLength(1);
    expect(all[0].label).toBe("Second");
  });

  it("trims whitespace on save", async () => {
    const a = await setAlias("conn-A", "project", 1, "  spaced  ");
    expect(a.label).toBe("spaced");
  });

  it("scopes lookups by connection id", async () => {
    await setAlias("conn-A", "trial", 7, "A-name");
    await setAlias("conn-B", "trial", 7, "B-name");
    expect((await getAlias("conn-A", "trial", 7))?.label).toBe("A-name");
    expect((await getAlias("conn-B", "trial", 7))?.label).toBe("B-name");
  });
});

describe("loadAliases", () => {
  it("returns only entries for the requested connection", async () => {
    await setAlias("conn-A", "trial", 1, "a1");
    await setAlias("conn-A", "objective", 2, "a2");
    await setAlias("conn-B", "project", 3, "b1");
    const aOnly = await loadAliases("conn-A");
    expect(aOnly).toHaveLength(2);
    expect(aOnly.every((a) => a.connectionId === "conn-A")).toBe(true);
  });
});

describe("clearAlias", () => {
  it("removes one alias and leaves others", async () => {
    await setAlias("conn-A", "trial", 1, "keep");
    await setAlias("conn-A", "objective", 2, "drop");
    await clearAlias("conn-A", "objective", 2);
    const remaining = await loadAliases("conn-A");
    expect(remaining).toHaveLength(1);
    expect(remaining[0].kind).toBe("trial");
  });

  it("is a no-op when nothing matches", async () => {
    await clearAlias("conn-X", "trial", 99);
    expect(await loadAllAliases()).toEqual([]);
  });
});

describe("setPinned", () => {
  it("creates a new pinned entry with empty label when none exists", async () => {
    const a = await setPinned("conn-A", "model", 7, true);
    expect(a).not.toBeNull();
    expect(a!.pinned).toBe(true);
    expect(a!.label).toBe("");
  });

  it("preserves label when pinning an existing aliased entry", async () => {
    await setAlias("conn-A", "project", 5, "My project");
    const a = await setPinned("conn-A", "project", 5, true);
    expect(a!.label).toBe("My project");
    expect(a!.pinned).toBe(true);
  });

  it("preserves pin when relabeling", async () => {
    await setPinned("conn-A", "trial", 9, true);
    await setAlias("conn-A", "trial", 9, "Renamed trial");
    const got = await getAlias("conn-A", "trial", 9);
    expect(got?.pinned).toBe(true);
    expect(got?.label).toBe("Renamed trial");
  });

  it("removes entry entirely when unpinning a label-less pin", async () => {
    await setPinned("conn-A", "model", 7, true);
    const res = await setPinned("conn-A", "model", 7, false);
    expect(res).toBeNull();
    expect(await loadAllAliases()).toHaveLength(0);
  });

  it("keeps entry but clears pin when unpinning an aliased+pinned entry", async () => {
    await setAlias("conn-A", "project", 1, "Keep me");
    await setPinned("conn-A", "project", 1, true);
    const res = await setPinned("conn-A", "project", 1, false);
    expect(res?.pinned).toBe(false);
    expect(res?.label).toBe("Keep me");
  });

  it("no-ops when unpinning a missing entry", async () => {
    const res = await setPinned("conn-A", "model", 999, false);
    expect(res).toBeNull();
    expect(await loadAllAliases()).toHaveLength(0);
  });
});

describe("clearAlias with pins", () => {
  it("keeps the entry as a label-less pin when pinned", async () => {
    await setAlias("conn-A", "model", 3, "Named");
    await setPinned("conn-A", "model", 3, true);
    await clearAlias("conn-A", "model", 3);
    const got = await getAlias("conn-A", "model", 3);
    expect(got).not.toBeNull();
    expect(got?.label).toBe("");
    expect(got?.pinned).toBe(true);
  });

  it("removes the entry when not pinned", async () => {
    await setAlias("conn-A", "model", 3, "Named");
    await clearAlias("conn-A", "model", 3);
    expect(await getAlias("conn-A", "model", 3)).toBeNull();
  });
});

describe("clearAliasesForConnection", () => {
  it("drops every alias scoped to the connection and preserves others", async () => {
    await setAlias("conn-A", "trial", 1, "a1");
    await setPinned("conn-A", "model", 2, true);
    await setAlias("conn-B", "project", 3, "b1");
    await clearAliasesForConnection("conn-A");
    const remaining = await loadAllAliases();
    expect(remaining).toHaveLength(1);
    expect(remaining[0].connectionId).toBe("conn-B");
  });

  it("is a no-op when nothing matches", async () => {
    await setAlias("conn-B", "trial", 1, "keep");
    await clearAliasesForConnection("conn-X");
    expect(await loadAllAliases()).toHaveLength(1);
  });
});

describe("isPinned", () => {
  it("returns true only for matching kind+id with pinned=true", async () => {
    await setPinned("conn-A", "model", 1, true);
    await setAlias("conn-A", "project", 1, "not pinned");
    const list = await loadAllAliases();
    expect(isPinned(list, "model", 1)).toBe(true);
    expect(isPinned(list, "project", 1)).toBe(false);
    expect(isPinned(list, "model", 2)).toBe(false);
  });
});

describe("displayName", () => {
  it("returns the alias label when present", async () => {
    const a = await setAlias("conn-A", "project", 5, "Pretty");
    expect(displayName([a], "project", 5, "Project.name")).toBe("Pretty");
  });

  it("returns the fallback when no alias is set", () => {
    expect(displayName([], "project", 5, "Project.name")).toBe("Project.name");
  });

  it("matches on (kind, id) and ignores other kinds with the same id", async () => {
    const a = await setAlias("conn-A", "trial", 5, "Trial five");
    expect(displayName([a], "project", 5, "Project default")).toBe("Project default");
    expect(displayName([a], "trial", 5, "Trial fallback")).toBe("Trial five");
  });
});
