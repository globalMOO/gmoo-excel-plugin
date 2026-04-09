import { useMemo } from "react";
import { GmooClient } from "../services/gmooApi";

const DEFAULT_GLOBALMOO_URL = "https://app.globalmoo.com/api/";
const PROXY_URL = "https://localhost:3001/api/";

/**
 * Resolves the API URL for the client.
 * If the user-facing URL is the default globalmoo.com, route through the local CORS proxy.
 * Otherwise, use the URL directly (assumes CORS is handled or it's a local server).
 */
function resolveBaseUrl(displayUrl: string): string {
  const normalized = displayUrl.trim().replace(/\/+$/, "") + "/";
  if (normalized === DEFAULT_GLOBALMOO_URL) {
    return PROXY_URL;
  }
  return normalized;
}

export function useGmooClient(apiKey: string, apiUrl: string): GmooClient | null {
  return useMemo(() => {
    if (!apiKey) return null;
    return new GmooClient(apiKey, resolveBaseUrl(apiUrl));
  }, [apiKey, apiUrl]);
}
