// React hook around aliasRegistryService, scoped to the active connection.
// Caches the list in state so the synchronous displayName() helper is cheap
// to call from render.

import { useCallback, useEffect, useState } from "react";
import {
  loadAliases,
  setAlias as svcSet,
  clearAlias as svcClear,
  setPinned as svcSetPinned,
  isPinned as svcIsPinned,
  displayName as svcDisplayName,
} from "../services/aliasRegistryService";
import type { EntityAlias, EntityKind } from "../types/aliasRegistry";

export interface UseAliasesResult {
  aliases: EntityAlias[];
  isLoading: boolean;
  /** Sync lookup; returns the alias label if set, else the fallback. */
  getName: (kind: EntityKind, id: number, fallback: string) => string;
  /** Persist a new label. Empty/whitespace label resets the alias instead. */
  rename: (kind: EntityKind, id: number, label: string) => Promise<void>;
  reset: (kind: EntityKind, id: number) => Promise<void>;
  /** True if the entity is currently pinned. */
  isPinned: (kind: EntityKind, id: number) => boolean;
  /** Toggle the pin flag; persists and refreshes state. */
  togglePin: (kind: EntityKind, id: number) => Promise<void>;
  /**
   * Stable comparator: pinned entries first (alphabetical by displayed name),
   * unpinned after (alphabetical). Picker option arrays pass this to .sort().
   */
  sortByPinned: <T>(
    items: T[],
    getKind: (item: T) => EntityKind,
    getId: (item: T) => number,
    getName: (item: T) => string
  ) => T[];
  refresh: () => Promise<void>;
}

export function useAliases(connectionId: string | null): UseAliasesResult {
  const [aliases, setAliases] = useState<EntityAlias[]>([]);
  const [isLoading, setIsLoading] = useState(true);

  const refresh = useCallback(async () => {
    if (!connectionId) {
      setAliases([]);
      setIsLoading(false);
      return;
    }
    const list = await loadAliases(connectionId);
    setAliases(list);
  }, [connectionId]);

  useEffect(() => {
    let cancelled = false;
    setIsLoading(true);
    (async () => {
      if (!connectionId) {
        if (!cancelled) {
          setAliases([]);
          setIsLoading(false);
        }
        return;
      }
      const list = await loadAliases(connectionId);
      if (!cancelled) {
        setAliases(list);
        setIsLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [connectionId]);

  const getName = useCallback(
    (kind: EntityKind, id: number, fallback: string) =>
      svcDisplayName(aliases, kind, id, fallback),
    [aliases]
  );

  const rename = useCallback(
    async (kind: EntityKind, id: number, label: string) => {
      if (!connectionId) return;
      const trimmed = label.trim();
      if (!trimmed) {
        await svcClear(connectionId, kind, id);
      } else {
        await svcSet(connectionId, kind, id, trimmed);
      }
      await refresh();
    },
    [connectionId, refresh]
  );

  const reset = useCallback(
    async (kind: EntityKind, id: number) => {
      if (!connectionId) return;
      await svcClear(connectionId, kind, id);
      await refresh();
    },
    [connectionId, refresh]
  );

  const isPinned = useCallback(
    (kind: EntityKind, id: number) => svcIsPinned(aliases, kind, id),
    [aliases]
  );

  const togglePin = useCallback(
    async (kind: EntityKind, id: number) => {
      if (!connectionId) return;
      const currentlyPinned = svcIsPinned(aliases, kind, id);
      await svcSetPinned(connectionId, kind, id, !currentlyPinned);
      await refresh();
    },
    [aliases, connectionId, refresh]
  );

  const sortByPinned = useCallback(
    <T,>(
      items: T[],
      getKind: (item: T) => EntityKind,
      getId: (item: T) => number,
      getDisplayName: (item: T) => string
    ): T[] => {
      return [...items].sort((a, b) => {
        const ap = svcIsPinned(aliases, getKind(a), getId(a));
        const bp = svcIsPinned(aliases, getKind(b), getId(b));
        if (ap !== bp) return ap ? -1 : 1;
        return getDisplayName(a).localeCompare(getDisplayName(b), undefined, {
          sensitivity: "base",
          numeric: true,
        });
      });
    },
    [aliases]
  );

  return {
    aliases,
    isLoading,
    getName,
    rename,
    reset,
    isPinned,
    togglePin,
    sortByPinned,
    refresh,
  };
}
