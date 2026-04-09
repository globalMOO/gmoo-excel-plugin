// TypeScript port of gmoo-sdk-csharp/ResponseTypes.cs and Enums.cs

// --- Enums ---

export enum ObjectiveType {
  Exact = "exact",
  Percent = "percent",
  Value = "value",
  LessThan = "lessthan",
  LessThanEqual = "lessthan_equal",
  GreaterThan = "greaterthan",
  GreaterThanEqual = "greaterthan_equal",
  Minimize = "minimize",
  Maximize = "maximize",
}

export enum InputType {
  Boolean = "boolean",
  Category = "category",
  Float = "float",
  Integer = "integer",
}

export enum StopReason {
  Running = 0,
  Satisfied = 1,
  Stopped = 2,
  Exhausted = 3,
}

// --- Response DTOs ---

export interface GmooBase {
  id: number;
  createdAt: string;
  updatedAt: string;
  disabledAt: string | null;
}

export interface Account extends GmooBase {
  firstName: string;
  lastName: string;
  company: string;
  name: string;
  email: string;
  apiKey: string;
  timeZone: string;
  customerId: string;
}

export interface GmooError {
  status: number;
  title: string;
  message: string;
  errors: Record<string, string>[];
}

export interface Model extends GmooBase {
  name: string;
  description: string;
  projects: Project[];
}

export interface Project extends GmooBase {
  developedAt: string | null;
  name: string;
  inputCount: number;
  minimums: number[];
  maximums: number[];
  inputTypes: string[];
  categories: string[];
  inputCases: number[][];
  caseCount: number;
  trials: Trial[];
}

export interface Trial extends GmooBase {
  number: number;
  outputCount: number;
  outputCases: number[][];
  caseCount: number;
  objectives: Objective[];
}

export interface Objective extends GmooBase {
  optimalInverse: Inverse | null;
  attemptCount: number;
  stopReason: number;
  desiredL1Norm: number;
  objectives: number[];
  objectiveTypes: string[];
  minimumBounds: number[];
  maximumBounds: number[];
  inverses: Inverse[];
}

export interface Inverse extends GmooBase {
  loadedAt: string | null;
  satisfiedAt: string | null;
  stoppedAt: string | null;
  exhaustedAt: string | null;
  iteration: number;
  l1Norm: number;
  suggestTime: number;
  computeTime: number;
  input: number[];
  output: number[];
  errors: number[];
  results: Result[];
}

export interface Result extends GmooBase {
  number: number;
  objective: number;
  objectiveType: string;
  minimumBound: number;
  maximumBound: number;
  output: number;
  error: number;
  detail: string;
  satisfied: boolean;
}

// --- Helper functions (mirrors C# SDK computed properties) ---

export function getStopReason(inverse: Inverse): StopReason {
  if (inverse.satisfiedAt) return StopReason.Satisfied;
  if (inverse.stoppedAt) return StopReason.Stopped;
  if (inverse.exhaustedAt) return StopReason.Exhausted;
  return StopReason.Running;
}

export function shouldStop(inverse: Inverse): boolean {
  return getStopReason(inverse) !== StopReason.Running;
}

export function getStopReasonLabel(reason: StopReason): string {
  switch (reason) {
    case StopReason.Satisfied:
      return "Objective satisfied";
    case StopReason.Stopped:
      return "Stopped (duplicate inputs suggested)";
    case StopReason.Exhausted:
      return "Exhausted all attempts";
    case StopReason.Running:
      return "Running";
  }
}

// --- L1 error helpers ---

/** Objective types that have a meaningful target error (Value, Percent, Exact) */
const TARGET_BASED_TYPES = new Set([
  ObjectiveType.Exact,
  ObjectiveType.Percent,
  ObjectiveType.Value,
]);

/** Whether this objective type contributes to target-based error */
export function isTargetBasedType(type: string): boolean {
  return TARGET_BASED_TYPES.has(type as ObjectiveType);
}

/**
 * Compute L1 norm using only target-based objectives (Value, Percent, Exact).
 * Inequality and min/max objectives are excluded since their "error" isn't meaningful.
 * Returns the API l1Norm if no results are available to filter.
 */
export function filteredL1Norm(inverse: Inverse): number {
  if (!inverse.results || inverse.results.length === 0) return inverse.l1Norm;
  const targetErrors = inverse.results.filter((r) => isTargetBasedType(r.objectiveType));
  if (targetErrors.length === 0) return 0;
  return targetErrors.reduce((sum, r) => sum + Math.abs(r.error), 0);
}

// --- Request DTOs ---

export interface CreateModelRequest {
  name: string;
  description?: string;
}

export interface CreateProjectRequest {
  name: string;
  inputCount: number;
  minimums: number[];
  maximums: number[];
  inputTypes: string[];
  categories?: string[];
}

export interface LoadOutputCasesRequest {
  outputCount: number;
  outputCases: number[][];
}

export interface LoadObjectivesRequest {
  desiredL1Norm: number;
  objectives: number[];
  objectiveTypes: string[];
  initialInput: number[];
  initialOutput: number[];
  minimumBounds?: number[];
  maximumBounds?: number[];
}

export interface LoadInverseOutputRequest {
  output: number[];
}
