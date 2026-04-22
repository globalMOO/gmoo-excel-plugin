import { useMemo } from "react";
import { GmooClient } from "../services/gmooApi";

export function useGmooClient(apiKey: string, apiUrl: string): GmooClient | null {
  return useMemo(() => {
    if (!apiKey) return null;
    const baseUrl = apiUrl.trim().replace(/\/+$/, "") + "/";
    return new GmooClient(apiKey, baseUrl);
  }, [apiKey, apiUrl]);
}
