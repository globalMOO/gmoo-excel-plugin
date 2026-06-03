// Random sampling of input vectors within each variable's bounds. Shared by
// Multi-Solve (random restarts) and the Set Objectives "randomize starting
// point" option so both draw points the same way.
import type { InputVariable } from "../types/workbookState";

/**
 * Draw one random input vector within the variables' bounds.
 *   • float            → uniform in [min, max]
 *   • integer/category → uniform integer in [floor(min), floor(max)] inclusive
 *   • boolean          → 0 or 1
 * Integer/category/boolean handling is retained for resumed legacy projects even
 * though the UI no longer creates those types.
 */
export function randomInputWithinBounds(variables: InputVariable[]): number[] {
  return variables.map((v) => {
    if (v.type === "boolean") return Math.random() < 0.5 ? 0 : 1;
    if (v.type === "integer" || v.type === "category") {
      const lo = Math.floor(v.min);
      const hi = Math.floor(v.max);
      if (hi <= lo) return lo;
      return lo + Math.floor(Math.random() * (hi - lo + 1));
    }
    return v.min + Math.random() * (v.max - v.min);
  });
}
