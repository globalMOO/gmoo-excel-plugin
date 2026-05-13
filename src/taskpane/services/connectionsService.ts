// Pure functions over the Connection list. No React.
// Persists to OfficeRuntime.storage under vsme.connections.v1, with a
// localStorage fallback (Office Web sometimes blocks OfficeRuntime.storage).
// All disk operations are read-modify-write the whole array — small list,
// atomicity beats efficiency.

import type { Connection, ConnectionSource, NewConnectionInput } from "../types/connection";

const STORAGE_KEY = "vsme.connections.v1";

// Pre-Connections storage keys (single apiKey/apiUrl pair). Migrated to a
// single Connection on first launch and then removed.
const LEGACY_KEY_API = "vsme_api_key";
const LEGACY_KEY_URL = "vsme_api_url";

const MAX_CONNECTIONS = 50;

// --- Storage primitives ---

async function loadRaw(key: string): Promise<string | null> {
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      const val = await OfficeRuntime.storage.getItem(key);
      if (val) return val;
    }
  } catch {
    // OfficeRuntime.storage unavailable
  }
  try {
    return localStorage.getItem(key);
  } catch {
    return null;
  }
}

async function saveRaw(key: string, value: string): Promise<void> {
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      await OfficeRuntime.storage.setItem(key, value);
    }
  } catch {
    // ignore
  }
  try {
    localStorage.setItem(key, value);
  } catch {
    // ignore
  }
}

async function removeRaw(key: string): Promise<void> {
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      await OfficeRuntime.storage.removeItem(key);
    }
  } catch {
    // ignore
  }
  try {
    localStorage.removeItem(key);
  } catch {
    // ignore
  }
}

// --- UUID v4 ---

function uuidv4(): string {
  // Prefer the platform implementation when available.
  const c =
    typeof globalThis !== "undefined"
      ? (globalThis as { crypto?: Crypto }).crypto
      : undefined;
  if (c && typeof c.randomUUID === "function") {
    return c.randomUUID();
  }
  // RFC 4122 v4 fallback using getRandomValues if present, else Math.random.
  const bytes = new Uint8Array(16);
  if (c && typeof c.getRandomValues === "function") {
    c.getRandomValues(bytes);
  } else {
    for (let i = 0; i < 16; i++) bytes[i] = Math.floor(Math.random() * 256);
  }
  bytes[6] = (bytes[6] & 0x0f) | 0x40; // version 4
  bytes[8] = (bytes[8] & 0x3f) | 0x80; // variant 10
  const hex: string[] = [];
  for (let i = 0; i < 16; i++) hex.push(bytes[i].toString(16).padStart(2, "0"));
  return (
    hex.slice(0, 4).join("") +
    "-" +
    hex.slice(4, 6).join("") +
    "-" +
    hex.slice(6, 8).join("") +
    "-" +
    hex.slice(8, 10).join("") +
    "-" +
    hex.slice(10, 16).join("")
  );
}

// --- Normalization ---

export function normalizeApiUrl(url: string): string {
  return url.trim().replace(/\/+$/, "");
}

/**
 * Strip a trailing /api or /api/ segment if present. Only used during legacy
 * migration: pre-Connections storage held the GmooClient base URL (which
 * includes /api/), but Connection.apiUrl is host-only — the /api/ suffix is
 * appended by callers that construct API URLs.
 */
function stripLegacyApiSuffix(url: string): string {
  return url.trim().replace(/\/+$/, "").replace(/\/api$/i, "");
}

function ensureUniqueLabel(desired: string, taken: Set<string>): string {
  if (!taken.has(desired)) return desired;
  for (let n = 2; n < 1000; n++) {
    const candidate = `${desired} (${n})`;
    if (!taken.has(candidate)) return candidate;
  }
  // Pathological — fall back to a UUID suffix.
  return `${desired} (${uuidv4().slice(0, 4)})`;
}

// --- Public API ---

export async function loadConnections(): Promise<Connection[]> {
  const raw = await loadRaw(STORAGE_KEY);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    // Tolerate forward-compatible extra fields; require the core ones.
    return parsed.filter(
      (c: unknown): c is Connection =>
        !!c &&
        typeof c === "object" &&
        typeof (c as Connection).id === "string" &&
        typeof (c as Connection).label === "string" &&
        typeof (c as Connection).apiUrl === "string" &&
        typeof (c as Connection).apiKey === "string"
    );
  } catch {
    return [];
  }
}

export async function saveConnections(list: Connection[]): Promise<void> {
  await saveRaw(STORAGE_KEY, JSON.stringify(list));
}

export async function createConnection(input: NewConnectionInput): Promise<Connection> {
  const list = await loadConnections();
  if (list.length >= MAX_CONNECTIONS) {
    throw new Error(`Maximum of ${MAX_CONNECTIONS} connections reached.`);
  }
  const taken = new Set(list.map((c) => c.label));
  const label = ensureUniqueLabel(input.label.trim() || "Connection", taken);
  const conn: Connection = {
    id: uuidv4(),
    label,
    apiUrl: normalizeApiUrl(input.apiUrl),
    apiKey: input.apiKey,
    source: input.source ?? "manual",
    createdAt: new Date().toISOString(),
  };
  await saveConnections([...list, conn]);
  return conn;
}

export async function updateConnection(
  id: string,
  patch: Partial<Connection>
): Promise<Connection> {
  const list = await loadConnections();
  const idx = list.findIndex((c) => c.id === id);
  if (idx === -1) throw new Error(`Connection not found: ${id}`);

  // Label uniqueness is enforced when label changes.
  let nextLabel = list[idx].label;
  if (patch.label !== undefined && patch.label !== list[idx].label) {
    const taken = new Set(list.filter((_, i) => i !== idx).map((c) => c.label));
    nextLabel = ensureUniqueLabel(patch.label.trim() || list[idx].label, taken);
  }

  const updated: Connection = {
    ...list[idx],
    ...patch,
    id: list[idx].id, // id is immutable
    label: nextLabel,
    apiUrl:
      patch.apiUrl !== undefined ? normalizeApiUrl(patch.apiUrl) : list[idx].apiUrl,
    createdAt: list[idx].createdAt, // createdAt is immutable
  };
  const next = [...list];
  next[idx] = updated;
  await saveConnections(next);
  return updated;
}

export async function deleteConnection(id: string): Promise<void> {
  const list = await loadConnections();
  await saveConnections(list.filter((c) => c.id !== id));
}

export async function findByUrlAndLabel(
  apiUrl: string,
  label: string
): Promise<Connection | null> {
  const target = normalizeApiUrl(apiUrl);
  const list = await loadConnections();
  return list.find((c) => c.apiUrl === target && c.label === label) ?? null;
}

export async function touchLastUsed(id: string): Promise<void> {
  await updateConnection(id, { lastUsedAt: new Date().toISOString() });
}

/**
 * If a pre-Connections install left an apiKey/apiUrl in storage, migrate it
 * into a single Connection labeled "Default (migrated)" and remove the legacy
 * keys. Returns the new connection, or null if there was nothing to migrate.
 *
 * Idempotent: if connections already exist OR the legacy keys are absent, no-op.
 */
export async function migrateLegacyKeyIfPresent(): Promise<Connection | null> {
  const existing = await loadConnections();
  if (existing.length > 0) {
    // Already migrated (or never had a legacy key). Clean up legacy keys
    // defensively in case a prior partial run left them around.
    await removeRaw(LEGACY_KEY_API);
    await removeRaw(LEGACY_KEY_URL);
    return null;
  }

  const legacyKey = await loadRaw(LEGACY_KEY_API);
  const legacyUrl = await loadRaw(LEGACY_KEY_URL);
  if (!legacyKey && !legacyUrl) return null;

  const conn = await createConnection({
    label: "Default (migrated)",
    apiUrl: stripLegacyApiSuffix(legacyUrl ?? "https://app.globalmoo.com"),
    apiKey: legacyKey ?? "",
    source: "manual",
  });
  await removeRaw(LEGACY_KEY_API);
  await removeRaw(LEGACY_KEY_URL);
  return conn;
}

// Re-export for callers that want the source-type union without a second import.
export type { Connection, ConnectionSource, NewConnectionInput };
