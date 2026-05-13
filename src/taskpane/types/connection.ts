// Connection model — replaces the singular apiKey/apiUrl pair.
// A Connection is the unit of "where do I send API calls and with what credentials."
// See Project-Spec.md §2 for the contract.

export type ConnectionSource =
  | "manual"      // User typed it in
  | "activation"  // Came from an activation token exchange
  | "admin-config" // Pushed by admin-deployed configuration (future)
  | "sso";        // Retrieved via SSO (future)

export interface Connection {
  id: string;             // UUID v4, generated client-side
  label: string;          // Human-readable, user-editable
  apiUrl: string;         // Base URL, no trailing slash
  apiKey: string;         // Opaque bearer token; may be empty
  source: ConnectionSource;
  createdAt: string;      // ISO 8601
  lastUsedAt?: string;    // ISO 8601
}

export interface NewConnectionInput {
  label: string;
  apiUrl: string;
  apiKey: string;
  source?: ConnectionSource; // defaults to "manual"
}
