// Fetches the model/project catalog for the active connection.
//
// The /api/models listing returns Models *without* their projects[] (the
// controller only applies the `read` serializer group; `read:projects` is
// gated behind /api/models/{id}). So to build a flat cross-model project
// list we have to fan out: getModels() then getModel(id) for each. There's
// no public endpoint that returns projects directly.

import { useCallback, useEffect, useState } from "react";
import type { GmooClient } from "../services/gmooApi";
import type { Model, Project } from "../types/gmoo";

export interface CatalogProject extends Project {
  modelId: number;
  modelName: string;
}

export interface UseProjectCatalogResult {
  models: Model[];
  /** Flat list of every project across all models, decorated with model info. */
  projects: CatalogProject[];
  isLoading: boolean;
  error: string | null;
  refresh: () => Promise<void>;
}

export function useProjectCatalog(client: GmooClient | null): UseProjectCatalogResult {
  const [models, setModels] = useState<Model[]>([]);
  const [projects, setProjects] = useState<CatalogProject[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const refresh = useCallback(async () => {
    if (!client) {
      setModels([]);
      setProjects([]);
      return;
    }
    setIsLoading(true);
    setError(null);
    try {
      const list = await client.getModels();
      setModels(list);
      // Fan out — /api/models gives us model metadata only, so a per-model
      // fetch is the only way to populate projects[]. Run in parallel; a
      // single failure shouldn't blank the whole catalog, so capture per-
      // model errors and keep going.
      const detailed = await Promise.all(
        list.map((m) =>
          client.getModel(m.id).catch(() => null)
        )
      );
      const flat: CatalogProject[] = [];
      for (let i = 0; i < list.length; i++) {
        const m = detailed[i] ?? list[i];
        for (const p of m.projects ?? []) {
          flat.push({ ...p, modelId: list[i].id, modelName: list[i].name });
        }
      }
      // Most-recently-updated projects first.
      flat.sort((a, b) => (a.updatedAt < b.updatedAt ? 1 : -1));
      setProjects(flat);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to load catalog.");
      setModels([]);
      setProjects([]);
    } finally {
      setIsLoading(false);
    }
  }, [client]);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  return { models, projects, isLoading, error, refresh };
}
