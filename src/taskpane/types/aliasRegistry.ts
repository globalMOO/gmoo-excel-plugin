// Local nicknames for server entities (project / trial / objective).
// Names live client-side because Trial and Objective have no name field on
// the API; Project does, but we still allow a local override so users can
// rename without round-tripping through the backend.
//
// Aliases are scoped per connection so the same id space doesn't collide
// across workspaces.

export type EntityKind = "model" | "project" | "trial" | "objective";

export interface EntityAlias {
  connectionId: string;
  kind: EntityKind;
  id: number;
  /** User-supplied rename. May be empty when the entry exists only to record `pinned`. */
  label: string;
  /** When true, UI surfaces this entry at the top of its picker. */
  pinned?: boolean;
  updatedAt: string;
}
