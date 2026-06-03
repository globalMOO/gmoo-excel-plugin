// Pure name-resolution helper for the Resume flow.
//
// The GMOO API does not store user-facing variable/outcome *names* — only
// counts, bounds, and types. Names live in workbook state (custom XML) and in
// the _GMOO_State sheet. When the user resumes a project whose names aren't in
// the current in-memory state (e.g. reopened on another machine, or after a
// connection switch wiped state), we must recover them from the _GMOO_State
// sheet rather than falling back to generic "Input N" / "Outcome N" labels —
// those generic labels otherwise leak into the template sheet and chart axes.
//
// Precedence per index:
//   1. same-project in-state name (the user may have just renamed it in-session)
//   2. matching name from the _GMOO_State sheet
//   3. generic fallback ("Input N" / "Outcome N")

function nonEmpty(s: string | undefined | null): string | undefined {
  if (s == null) return undefined;
  const t = String(s).trim();
  return t.length > 0 ? t : undefined;
}

export interface ResolveResumeNamesParams {
  inputCount: number;
  outputCount: number;
  /** True when resuming the same project that's currently in state. */
  sameProject: boolean;
  /** In-memory state names (only trusted when sameProject). */
  stateInputNames: (string | undefined)[];
  stateOutcomeNames: (string | undefined)[];
  /** Names recovered from the _GMOO_State sheet. Pass only when they describe
   *  the project being resumed (caller verifies the projectId tag + count). */
  savedVariableNames?: string[];
  savedOutcomeNames?: string[];
}

export function resolveResumeNames(
  params: ResolveResumeNamesParams
): { inputNames: string[]; outcomeNames: string[] } {
  const {
    inputCount,
    outputCount,
    sameProject,
    stateInputNames,
    stateOutcomeNames,
    savedVariableNames,
    savedOutcomeNames,
  } = params;

  const inputNames: string[] = [];
  for (let i = 0; i < inputCount; i++) {
    const stateName = sameProject ? nonEmpty(stateInputNames[i]) : undefined;
    const savedName = nonEmpty(savedVariableNames?.[i]);
    inputNames.push(stateName ?? savedName ?? `Input ${i + 1}`);
  }

  const outcomeNames: string[] = [];
  for (let i = 0; i < outputCount; i++) {
    const stateName = sameProject ? nonEmpty(stateOutcomeNames[i]) : undefined;
    const savedName = nonEmpty(savedOutcomeNames?.[i]);
    outcomeNames.push(stateName ?? savedName ?? `Outcome ${i + 1}`);
  }

  return { inputNames, outcomeNames };
}
