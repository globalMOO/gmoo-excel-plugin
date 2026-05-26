// Pure functions over the EntityAlias list. No React.
// Mirrors connectionsService: read-modify-write the whole array, persisted
// to OfficeRuntime.storage with a localStorage fallback.

import type { EntityAlias, EntityKind } from "../types/aliasRegistry";

const STORAGE_KEY = "gmoo.aliases.v1";

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

function isAlias(c: unknown): c is EntityAlias {
  if (!c || typeof c !== "object") return false;
  const a = c as EntityAlias;
  return (
    typeof a.connectionId === "string" &&
    (a.kind === "model" || a.kind === "project" || a.kind === "trial" || a.kind === "objective") &&
    typeof a.id === "number" &&
    typeof a.label === "string"
  );
}

export async function loadAllAliases(): Promise<EntityAlias[]> {
  const raw = await loadRaw(STORAGE_KEY);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.filter(isAlias);
  } catch {
    return [];
  }
}

async function saveAll(list: EntityAlias[]): Promise<void> {
  await saveRaw(STORAGE_KEY, JSON.stringify(list));
}

export async function loadAliases(connectionId: string): Promise<EntityAlias[]> {
  const all = await loadAllAliases();
  return all.filter((a) => a.connectionId === connectionId);
}

/**
 * Remove every alias scoped to the given connection. Called when a connection
 * is deleted so stale pins/renames don't reappear if the same numeric ids show
 * up under a different connection later.
 */
export async function clearAliasesForConnection(connectionId: string): Promise<void> {
  const all = await loadAllAliases();
  const next = all.filter((a) => a.connectionId !== connectionId);
  if (next.length !== all.length) {
    await saveAll(next);
  }
}

export async function getAlias(
  connectionId: string,
  kind: EntityKind,
  id: number
): Promise<EntityAlias | null> {
  const all = await loadAllAliases();
  return (
    all.find(
      (a) => a.connectionId === connectionId && a.kind === kind && a.id === id
    ) ?? null
  );
}

export async function setAlias(
  connectionId: string,
  kind: EntityKind,
  id: number,
  label: string
): Promise<EntityAlias> {
  const trimmed = label.trim();
  const all = await loadAllAliases();
  const idx = all.findIndex(
    (a) => a.connectionId === connectionId && a.kind === kind && a.id === id
  );
  const now = new Date().toISOString();
  // Preserve pin status if an existing entry is being relabeled.
  const pinned = idx === -1 ? undefined : all[idx].pinned;
  const next: EntityAlias = { connectionId, kind, id, label: trimmed, pinned, updatedAt: now };
  if (idx === -1) {
    all.push(next);
  } else {
    all[idx] = next;
  }
  await saveAll(all);
  return next;
}

/**
 * Clear the rename only. If the entry is still pinned, it stays — just with
 * an empty label so the picker falls back to the canonical name. If neither
 * label nor pin survives, the entry is removed entirely.
 */
export async function clearAlias(
  connectionId: string,
  kind: EntityKind,
  id: number
): Promise<void> {
  const all = await loadAllAliases();
  const idx = all.findIndex(
    (a) => a.connectionId === connectionId && a.kind === kind && a.id === id
  );
  if (idx === -1) return;
  const entry = all[idx];
  if (entry.pinned) {
    all[idx] = { ...entry, label: "", updatedAt: new Date().toISOString() };
  } else {
    all.splice(idx, 1);
  }
  await saveAll(all);
}

/**
 * Set or clear the pin flag. Creates the entry if missing (with empty label);
 * removes the entry entirely if unpinning and no label is set.
 */
export async function setPinned(
  connectionId: string,
  kind: EntityKind,
  id: number,
  pinned: boolean
): Promise<EntityAlias | null> {
  const all = await loadAllAliases();
  const idx = all.findIndex(
    (a) => a.connectionId === connectionId && a.kind === kind && a.id === id
  );
  const now = new Date().toISOString();
  if (idx === -1) {
    if (!pinned) return null; // nothing to remove
    const created: EntityAlias = {
      connectionId,
      kind,
      id,
      label: "",
      pinned: true,
      updatedAt: now,
    };
    all.push(created);
    await saveAll(all);
    return created;
  }
  const entry = all[idx];
  if (!pinned && !entry.label) {
    all.splice(idx, 1);
    await saveAll(all);
    return null;
  }
  const updated: EntityAlias = { ...entry, pinned, updatedAt: now };
  all[idx] = updated;
  await saveAll(all);
  return updated;
}

export function isPinned(aliases: EntityAlias[], kind: EntityKind, id: number): boolean {
  return aliases.some((a) => a.kind === kind && a.id === id && !!a.pinned);
}

/**
 * Synchronous lookup over a pre-loaded list. UI components hold the list in
 * state via useAliases() and call this on every render, so we keep it sync.
 */
export function displayName(
  aliases: EntityAlias[],
  kind: EntityKind,
  id: number,
  fallback: string
): string {
  const match = aliases.find((a) => a.kind === kind && a.id === id);
  if (match && match.label) return match.label;
  return fallback;
}
