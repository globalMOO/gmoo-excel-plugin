import { useMemo } from "react";
import { GmooClient } from "../services/gmooApi";
import type { Connection } from "../types/connection";

const PROD_HOST = "https://app.globalmoo.com";

/**
 * In dev (taskpane served from https://localhost:3000), the prod API's CORS
 * whitelist rejects our origin. Webpack's dev server is configured to proxy
 * /api/* to https://app.globalmoo.com, so we rewrite the GmooClient base URL
 * to be same-origin and let the proxy do the cross-origin call server-to-server.
 * In a real build, window.location.hostname isn't "localhost", so this no-ops.
 */
function resolveHost(connectionHost: string): string {
  if (
    typeof window !== "undefined" &&
    window.location.hostname === "localhost" &&
    connectionHost === PROD_HOST
  ) {
    return `${window.location.protocol}//${window.location.host}`;
  }
  return connectionHost;
}

/**
 * Build a GmooClient from the active Connection. Returns null if there's no
 * active connection, or the connection has no apiKey yet (e.g. an
 * activation-seeded placeholder waiting on the user to finish a flow).
 *
 * Connection.apiUrl is host-only (e.g. "https://app.globalmoo.com"). The
 * GmooClient base URL needs the /api/ path, so we append it here.
 */
export function useGmooClient(connection: Connection | null): GmooClient | null {
  return useMemo(() => {
    if (!connection || !connection.apiKey) return null;
    // Tolerate users pasting a URL that already ends with /api or /api/ — the
    // spec says Connection.apiUrl is host-only, but the legacy stored value
    // included the path and copy-paste happens.
    const host = connection.apiUrl.replace(/\/+$/, "").replace(/\/api$/i, "");
    const base = resolveHost(host) + "/api/";
    return new GmooClient(connection.apiKey, base);
  }, [connection?.apiKey, connection?.apiUrl]);
}
