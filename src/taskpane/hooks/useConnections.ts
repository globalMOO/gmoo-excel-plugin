// React hook around connectionsService. Owns:
//   - the connection list (loaded once on mount, kept in sync after mutations)
//   - the active-connection state, derived from the workbook state field
//     activeConnectionId. Falls back to most-recently-used when the active id
//     points at a deleted connection.
//
// Runs migrateLegacyKeyIfPresent on first mount so users coming from the
// pre-Connections add-in don't have to re-enter their key.

import { useCallback, useEffect, useState } from "react";
import {
  loadConnections,
  createConnection as svcCreate,
  updateConnection as svcUpdate,
  deleteConnection as svcDelete,
  migrateLegacyKeyIfPresent,
  touchLastUsed,
} from "../services/connectionsService";
import type { Connection, NewConnectionInput } from "../types/connection";

export interface UseConnectionsOptions {
  /** Currently-active connection id (sourced from workbook state). */
  activeConnectionId: string | null;
  /** Persists the active connection id to workbook state. */
  setActiveConnectionId: (id: string | null) => Promise<void> | void;
}

export interface UseConnectionsResult {
  connections: Connection[];
  activeConnection: Connection | null;
  isLoading: boolean;
  setActive: (id: string) => Promise<void>;
  add: (input: NewConnectionInput) => Promise<Connection>;
  update: (id: string, patch: Partial<Connection>) => Promise<Connection>;
  remove: (id: string) => Promise<void>;
  /** Force-reload the list (used after activation creates/updates a connection externally). */
  refresh: () => Promise<void>;
}

/**
 * Pick the connection to treat as active given the user's stored preference.
 * Falls back to most-recently-used (lastUsedAt → createdAt) when the stored
 * id is missing or stale.
 */
function pickActive(list: Connection[], preferredId: string | null): Connection | null {
  if (list.length === 0) return null;
  if (preferredId) {
    const match = list.find((c) => c.id === preferredId);
    if (match) return match;
  }
  // Most-recently-used wins. lastUsedAt > createdAt > insertion order.
  return [...list].sort((a, b) => {
    const aTime = a.lastUsedAt ?? a.createdAt;
    const bTime = b.lastUsedAt ?? b.createdAt;
    if (aTime === bTime) return 0;
    return aTime > bTime ? -1 : 1;
  })[0];
}

export function useConnections(opts: UseConnectionsOptions): UseConnectionsResult {
  const { activeConnectionId, setActiveConnectionId } = opts;
  const [connections, setConnections] = useState<Connection[]>([]);
  const [isLoading, setIsLoading] = useState(true);

  const refresh = useCallback(async () => {
    const list = await loadConnections();
    setConnections(list);
  }, []);

  // First-mount: run legacy migration, then load.
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        await migrateLegacyKeyIfPresent();
      } catch {
        // Migration failure shouldn't block the UI — user can still add a
        // connection manually.
      }
      const list = await loadConnections();
      if (!cancelled) {
        setConnections(list);
        setIsLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  const activeConnection = pickActive(connections, activeConnectionId);

  // If the stored activeConnectionId doesn't match an existing connection but
  // we did pick a fallback, write the fallback back to workbook state so the
  // next render is in sync. Skip during initial load to avoid clobbering a
  // valid id before the connection list arrives.
  useEffect(() => {
    if (isLoading) return;
    if (!activeConnection) {
      if (activeConnectionId !== null) {
        // Stored id is stale and there's no fallback (empty list).
        void setActiveConnectionId(null);
      }
      return;
    }
    if (activeConnection.id !== activeConnectionId) {
      void setActiveConnectionId(activeConnection.id);
    }
  }, [isLoading, activeConnection, activeConnectionId, setActiveConnectionId]);

  const setActive = useCallback(
    async (id: string) => {
      await setActiveConnectionId(id);
      // Bump lastUsedAt so MRU fallback reflects the user's choice.
      try {
        await touchLastUsed(id);
        await refresh();
      } catch {
        // touchLastUsed throws if the id was just deleted; tolerate.
      }
    },
    [refresh, setActiveConnectionId]
  );

  const add = useCallback(
    async (input: NewConnectionInput) => {
      const conn = await svcCreate(input);
      await refresh();
      return conn;
    },
    [refresh]
  );

  const update = useCallback(
    async (id: string, patch: Partial<Connection>) => {
      const conn = await svcUpdate(id, patch);
      await refresh();
      return conn;
    },
    [refresh]
  );

  const remove = useCallback(
    async (id: string) => {
      await svcDelete(id);
      // If we just deleted the active one, clear the workbook pointer; the
      // sync effect above will pick a fallback on the next render.
      if (id === activeConnectionId) {
        await setActiveConnectionId(null);
      }
      await refresh();
    },
    [activeConnectionId, refresh, setActiveConnectionId]
  );

  return {
    connections,
    activeConnection,
    isLoading,
    setActive,
    add,
    update,
    remove,
    refresh,
  };
}
