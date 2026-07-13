// Tests stateStore against a strict Office.js fake that mirrors the real
// proxy-loading rules: reading a property that was not load()-ed before the
// last context.sync() throws, and getItemOrNullObject returns a fresh
// unloaded proxy on every call. This is what makes the fallback-path
// regression (reading re-fetched, never-synced proxies) actually fail here.
import { saveWorkbookState, loadWorkbookState, clearWorkbookState } from "../stateStore";
import { DEFAULT_WORKBOOK_STATE, STATE_SCHEMA_VERSION, WorkbookState } from "../../types/workbookState";

const XML_NAMESPACE = "gmoo-globalmoo-state";

interface FakeHostStores {
  xmlParts: { namespace: string; xml: string }[];
  settings: Map<string, string>;
}

function notLoaded(prop: string): Error {
  return new Error(
    `PropertyNotLoaded: The property '${prop}' is not available. Before reading the property's value, call the load method.`
  );
}

class FakeClientObject {
  protected loadedProps = new Set<string>();
  private pendingProps = new Set<string>();

  constructor(private ctx: FakeContext) {}

  load(props: string): void {
    for (const p of props.split(",").map((s) => s.trim()).filter(Boolean)) {
      this.pendingProps.add(p);
    }
    this.ctx.registerPending(this);
  }

  applySync(): void {
    for (const p of this.pendingProps) this.loadedProps.add(p);
    this.pendingProps.clear();
    this.populate();
  }

  protected populate(): void {}

  protected read<T>(prop: string, value: T): T {
    if (!this.loadedProps.has(prop)) throw notLoaded(prop);
    return value;
  }
}

class FakeSetting extends FakeClientObject {
  private snapshot: string | undefined;

  constructor(ctx: FakeContext, private store: Map<string, string>, private key: string) {
    super(ctx);
  }

  protected populate(): void {
    this.snapshot = this.store.get(this.key);
  }

  get isNullObject(): boolean {
    return this.read("isNullObject", this.snapshot === undefined);
  }

  get value(): string {
    return this.read("value", this.snapshot as string);
  }

  delete(): void {
    this.store.delete(this.key);
  }
}

class FakeXmlPart {
  constructor(
    private ctx: FakeContext,
    private stores: FakeHostStores,
    private entry: { namespace: string; xml: string }
  ) {}

  getXml(): { value: string } {
    const result = { value: undefined as unknown as string };
    this.ctx.onSync(() => {
      result.value = this.entry.xml;
    });
    return result;
  }

  delete(): void {
    const idx = this.stores.xmlParts.indexOf(this.entry);
    if (idx !== -1) this.stores.xmlParts.splice(idx, 1);
  }
}

class FakeXmlPartCollection extends FakeClientObject {
  private snapshot: FakeXmlPart[] = [];

  constructor(private ctx2: FakeContext, private stores: FakeHostStores, private namespace: string) {
    super(ctx2);
  }

  protected populate(): void {
    this.snapshot = this.stores.xmlParts
      .filter((p) => p.namespace === this.namespace)
      .map((p) => new FakeXmlPart(this.ctx2, this.stores, p));
  }

  get items(): FakeXmlPart[] {
    return this.read("items", this.snapshot);
  }

  add(xml: string): void {
    this.stores.xmlParts.push({ namespace: this.namespace, xml });
  }
}

class FakeContext {
  private pending: FakeClientObject[] = [];
  private syncCallbacks: (() => void)[] = [];

  workbook = {
    customXmlParts: {
      getByNamespace: (ns: string) => new FakeXmlPartCollection(this, this.stores, ns),
      add: (xml: string) => {
        // The add-in only writes parts in its own namespace
        this.stores.xmlParts.push({ namespace: XML_NAMESPACE, xml });
      },
    },
    settings: {
      getItemOrNullObject: (key: string) => new FakeSetting(this, this.stores.settings, key),
      add: (key: string, value: string) => {
        this.stores.settings.set(key, value);
      },
    },
  };

  constructor(private stores: FakeHostStores) {}

  registerPending(obj: FakeClientObject): void {
    this.pending.push(obj);
  }

  onSync(cb: () => void): void {
    this.syncCallbacks.push(cb);
  }

  async sync(): Promise<void> {
    for (const obj of this.pending) obj.applySync();
    this.pending = [];
    for (const cb of this.syncCallbacks) cb();
    this.syncCallbacks = [];
  }
}

let stores: FakeHostStores;

beforeEach(() => {
  stores = { xmlParts: [], settings: new Map() };
  (global as Record<string, unknown>).Excel = {
    run: async (cb: (ctx: FakeContext) => Promise<unknown>) => cb(new FakeContext(stores)),
  };
});

const sampleState: WorkbookState = {
  ...DEFAULT_WORKBOOK_STATE,
  apiKeyHint: "...key9",
  activeConnectionId: "conn-1",
  modelId: 42,
  modelName: "Solar Twin",
  projectId: 7,
  trialId: 12,
  objectiveId: 99,
  outcomeNames: ["Cost", "Yield"],
  wizardStep: 5,
};

describe("saveWorkbookState / loadWorkbookState round-trip", () => {
  it("restores the full state from the custom XML part", async () => {
    await saveWorkbookState(sampleState);
    const loaded = await loadWorkbookState();
    expect(loaded).toEqual(sampleState);
  });

  it("replaces the existing XML part instead of accumulating parts", async () => {
    await saveWorkbookState(sampleState);
    await saveWorkbookState({ ...sampleState, modelId: 43 });
    expect(stores.xmlParts).toHaveLength(1);
    const loaded = await loadWorkbookState();
    expect(loaded.modelId).toBe(43);
  });

  it("round-trips XML-significant characters in state values", async () => {
    const tricky = { ...sampleState, modelName: `<b>&"it's"</b>` };
    await saveWorkbookState(tricky);
    const loaded = await loadWorkbookState();
    expect(loaded.modelName).toBe(`<b>&"it's"</b>`);
  });
});

describe("loadWorkbookState settings fallback", () => {
  it("restores critical IDs from settings when no XML part exists", async () => {
    await saveWorkbookState(sampleState);
    stores.xmlParts = []; // simulate the XML part being lost

    const loaded = await loadWorkbookState();
    expect(loaded.modelId).toBe(42);
    expect(loaded.projectId).toBe(7);
    expect(loaded.trialId).toBe(12);
    expect(loaded.objectiveId).toBe(99);
    expect(loaded.wizardStep).toBe(5);
    expect(loaded.apiKeyHint).toBe("...key9");
    expect(loaded.activeConnectionId).toBe("conn-1");
    // Non-fallback fields come back as defaults
    expect(loaded.modelName).toBe("");
  });

  it("falls back to settings when the XML part is corrupted", async () => {
    await saveWorkbookState(sampleState);
    stores.xmlParts[0].xml = `<gmooState xmlns="${XML_NAMESPACE}">{not json</gmooState>`;

    const loaded = await loadWorkbookState();
    expect(loaded.modelId).toBe(42);
    expect(loaded.wizardStep).toBe(5);
  });

  it("applies the v1→v2 wizardStep migration to legacy settings without a schemaVersion", async () => {
    stores.settings.set("gmoo_modelId", "42");
    stores.settings.set("gmoo_wizardStep", "3");

    const loaded = await loadWorkbookState();
    expect(loaded.wizardStep).toBe(4);
    expect(loaded.schemaVersion).toBe(STATE_SCHEMA_VERSION);
  });

  it("does not re-migrate settings already at the current schema version", async () => {
    stores.settings.set("gmoo_schemaVersion", String(STATE_SCHEMA_VERSION));
    stores.settings.set("gmoo_wizardStep", "3");

    const loaded = await loadWorkbookState();
    expect(loaded.wizardStep).toBe(3);
  });

  it("returns defaults for an empty workbook", async () => {
    const loaded = await loadWorkbookState();
    expect(loaded).toEqual(DEFAULT_WORKBOOK_STATE);
  });
});

describe("clearWorkbookState", () => {
  it("removes the XML part and all fallback settings", async () => {
    await saveWorkbookState(sampleState);
    await clearWorkbookState();

    expect(stores.xmlParts).toHaveLength(0);
    expect(stores.settings.size).toBe(0);

    const loaded = await loadWorkbookState();
    expect(loaded).toEqual(DEFAULT_WORKBOOK_STATE);
  });

  it("is a no-op on an empty workbook", async () => {
    await expect(clearWorkbookState()).resolves.toBeUndefined();
  });
});
