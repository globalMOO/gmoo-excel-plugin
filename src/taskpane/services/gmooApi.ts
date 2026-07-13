// TypeScript port of gmoo-sdk-csharp/Client.cs
// Uses fetch() with Bearer token auth, retry with exponential backoff

import type {
  Model,
  Project,
  Trial,
  Objective,
  Inverse,
  GmooError,
  CreateModelRequest,
  CreateProjectRequest,
  LoadOutputCasesRequest,
  LoadObjectivesRequest,
  LoadInverseOutputRequest,
} from "../types/gmoo";

// Endpoints actually exposed by gmoo-api (see config/routes.yaml under "# API"):
//   GET    /api/models                              → list w/o projects[]
//   POST   /api/models
//   GET    /api/models/{id}                         → full tree (projects → trials → objectives → inverses)
//   POST   /api/models/{id}/projects
//   POST   /api/projects/{id}/output-cases
//   POST   /api/trials/{id}/objectives
//   GET    /api/objectives/{id}                     → objective w/ inverses[] (incl. results[])
//   POST   /api/objectives/{id}/suggest-inverse
//   POST   /api/inverses/{id}/load-output
// There is no GET /api/projects/{id}, /api/trials/{id}, or /api/inverses/{id}.
// All project/trial lookups must walk the model tree from getModel().

const VALID_INPUT_TYPES = ["boolean", "category", "float", "integer"];
const MAX_RETRIES = 3;
const DEFAULT_BASE_URL = "https://app.globalmoo.com/api/";
// suggest-inverse runs the optimizer server-side and can legitimately take a
// while on large models, so the default is generous.
const DEFAULT_TIMEOUT_MS = 60_000;

export class GmooApiError extends Error {
  constructor(
    public status: number,
    public apiError: GmooError | null,
    message?: string
  ) {
    super(message ?? `API error ${status}: ${apiError?.message ?? "Unknown error"}`);
    this.name = "GmooApiError";
  }
}

/** A request exceeded the client's timeout. Not retried. */
export class GmooTimeoutError extends Error {
  constructor(timeoutMs: number) {
    super(`Request timed out after ${Math.round(timeoutMs / 1000)}s.`);
    this.name = "GmooTimeoutError";
  }
}

/** A request was cancelled via the caller's AbortSignal. Not retried. */
export class GmooCancelledError extends Error {
  constructor() {
    super("Request cancelled.");
    this.name = "GmooCancelledError";
  }
}

export interface GmooClientOptions {
  /** Per-request timeout in milliseconds. 0 disables the timeout. */
  timeoutMs?: number;
}

export class GmooClient {
  private apiKey: string;
  private baseUrl: string;
  private timeoutMs: number;

  constructor(apiKey: string, baseUrl: string = DEFAULT_BASE_URL, options?: GmooClientOptions) {
    if (!apiKey) {
      throw new Error("API key cannot be empty.");
    }
    this.apiKey = apiKey;
    this.baseUrl = baseUrl.endsWith("/") ? baseUrl : baseUrl + "/";
    this.timeoutMs = options?.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  }

  // --- Model Operations ---

  async getModels(signal?: AbortSignal): Promise<Model[]> {
    return this.get<Model[]>("models", signal);
  }

  async getModel(modelId: number, signal?: AbortSignal): Promise<Model> {
    if (modelId <= 0) throw new Error("Model ID must be greater than zero.");
    return this.get<Model>(`models/${modelId}`, signal);
  }

  async createModel(name: string, description?: string, signal?: AbortSignal): Promise<Model> {
    if (!name || !name.trim()) throw new Error("Model name cannot be empty.");
    const request: CreateModelRequest = { name, description };
    return this.post<Model>("models", request, signal);
  }

  // --- Project Operations ---

  async createProject(
    modelId: number,
    name: string,
    inputCount: number,
    minimums: number[],
    maximums: number[],
    inputTypes: string[],
    categories?: string[],
    signal?: AbortSignal
  ): Promise<Project> {
    // Validation matching C# SDK
    if (modelId <= 0) throw new Error("Model ID must be greater than zero.");
    if (!name || name.trim().length < 4)
      throw new Error("Project name must be at least 4 characters long.");
    if (inputCount < 2 || inputCount > 200)
      throw new Error("Input count must be between 2 and 200.");
    if (minimums.length !== inputCount)
      throw new Error(`Length of minimums (${minimums.length}) does not match input count (${inputCount}).`);
    if (maximums.length !== inputCount)
      throw new Error(`Length of maximums (${maximums.length}) does not match input count (${inputCount}).`);
    if (inputTypes.length !== inputCount)
      throw new Error(`Length of inputTypes (${inputTypes.length}) does not match input count (${inputCount}).`);

    for (const t of inputTypes) {
      if (!VALID_INPUT_TYPES.includes(t.toLowerCase())) {
        throw new Error(`Invalid input type: ${t}. Valid types are: ${VALID_INPUT_TYPES.join(", ")}`);
      }
    }

    // Validate mins < maxs for non-boolean, non-category types
    for (let i = 0; i < inputCount; i++) {
      const t = inputTypes[i].toLowerCase();
      if (t !== "boolean" && t !== "category" && minimums[i] >= maximums[i]) {
        throw new Error(`Minimum (${minimums[i]}) must be less than maximum (${maximums[i]}) for input ${i}.`);
      }
    }

    const request: CreateProjectRequest = {
      name,
      inputCount,
      minimums,
      maximums,
      inputTypes,
      categories: categories ?? [],
    };
    return this.post<Project>(`models/${modelId}/projects`, request, signal);
  }

  // --- Output Cases ---

  async loadOutputCases(
    projectId: number,
    outputCount: number,
    outputCases: number[][],
    signal?: AbortSignal
  ): Promise<Trial> {
    if (projectId <= 0) throw new Error("Project ID must be greater than zero.");
    if (outputCount <= 0) throw new Error("Output count must be greater than zero.");
    if (!Array.isArray(outputCases) || outputCases.length === 0) {
      throw new Error("Output cases must be a non-empty matrix.");
    }
    for (let i = 0; i < outputCases.length; i++) {
      const oc = outputCases[i];
      if (!Array.isArray(oc) || oc.length !== outputCount) {
        throw new Error(`Output case ${i + 1} must be an array of length ${outputCount} (got ${Array.isArray(oc) ? oc.length : typeof oc}).`);
      }
      for (let j = 0; j < oc.length; j++) {
        if (typeof oc[j] !== "number" || !Number.isFinite(oc[j])) {
          throw new Error(`Output case ${i + 1}, value ${j + 1} is not a finite number (got ${oc[j]}). Check your Excel formulas for empty cells or errors.`);
        }
      }
    }
    const request: LoadOutputCasesRequest = { outputCount, outputCases };
    return this.post<Trial>(`projects/${projectId}/output-cases`, request, signal);
  }

  // --- Objective Operations ---

  async getObjective(objectiveId: number, signal?: AbortSignal): Promise<Objective> {
    if (objectiveId <= 0) throw new Error("Objective ID must be greater than zero.");
    return this.get<Objective>(`objectives/${objectiveId}`, signal);
  }

  async loadObjectives(
    trialId: number,
    objectives: number[],
    objectiveTypes: string[],
    initialInput: number[],
    initialOutput: number[],
    desiredL1Norm: number = 0.0,
    minimumBounds?: number[],
    maximumBounds?: number[],
    signal?: AbortSignal
  ): Promise<Objective> {
    if (trialId <= 0) throw new Error("Trial ID must be greater than zero.");
    if (objectives.length !== objectiveTypes.length)
      throw new Error("Number of objectives must match number of objective types.");

    // Default bounds for exact type (matching C# SDK behavior)
    let minBounds = minimumBounds;
    let maxBounds = maximumBounds;
    if (objectiveTypes.length > 0 && objectiveTypes[0] === "exact") {
      minBounds ??= new Array(objectives.length).fill(0);
      maxBounds ??= new Array(objectives.length).fill(0);
    }

    const request: LoadObjectivesRequest = {
      desiredL1Norm,
      objectives,
      objectiveTypes,
      initialInput,
      initialOutput,
      minimumBounds: minBounds,
      maximumBounds: maxBounds,
    };
    return this.post<Objective>(`trials/${trialId}/objectives`, request, signal);
  }

  // --- Inverse Operations ---

  async suggestInverse(objectiveId: number, signal?: AbortSignal): Promise<Inverse> {
    if (objectiveId <= 0) throw new Error("Objective ID must be greater than zero.");
    return this.post<Inverse>(`objectives/${objectiveId}/suggest-inverse`, {}, signal);
  }

  async loadInverseOutput(inverseId: number, output: number[], signal?: AbortSignal): Promise<Inverse> {
    if (inverseId <= 0) throw new Error("Inverse ID must be greater than zero.");
    if (output.length === 0) throw new Error("Output list cannot be empty.");
    const request: LoadInverseOutputRequest = { output };
    return this.post<Inverse>(`inverses/${inverseId}/load-output`, request, signal);
  }

  // --- HTTP Helpers ---

  private async get<T>(endpoint: string, signal?: AbortSignal): Promise<T> {
    return this.sendRequest<T>("GET", endpoint, undefined, signal);
  }

  private async post<T>(endpoint: string, data: unknown, signal?: AbortSignal): Promise<T> {
    return this.sendRequest<T>("POST", endpoint, data, signal);
  }

  private async sendRequest<T>(
    method: string,
    endpoint: string,
    data?: unknown,
    signal?: AbortSignal
  ): Promise<T> {
    let lastError: Error | null = null;

    for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
      try {
        return await this.executeRequest<T>(method, endpoint, data, signal);
      } catch (err) {
        lastError = err as Error;
        if (err instanceof GmooApiError && this.shouldRetry(err.status)) {
          if (attempt < MAX_RETRIES) {
            const waitTime = Math.min(4 * Math.pow(2, attempt), 10) * 1000;
            await this.delay(waitTime);
            continue;
          }
        }
        throw err;
      }
    }

    throw lastError ?? new Error("Request failed after retries");
  }

  private async executeRequest<T>(
    method: string,
    endpoint: string,
    data?: unknown,
    signal?: AbortSignal
  ): Promise<T> {
    if (signal?.aborted) throw new GmooCancelledError();

    const url = this.baseUrl + endpoint;
    const headers: Record<string, string> = {
      Authorization: `Bearer ${this.apiKey}`,
      Accept: "application/json",
    };

    // One controller drives the fetch; it fires on timeout or when the
    // caller's signal aborts (AbortSignal.any isn't available in all hosts).
    const controller = new AbortController();
    const timeoutId =
      this.timeoutMs > 0 ? setTimeout(() => controller.abort(), this.timeoutMs) : null;
    const onCallerAbort = () => controller.abort();
    signal?.addEventListener("abort", onCallerAbort);

    const init: RequestInit = { method, headers, signal: controller.signal };

    if (data && method !== "GET") {
      headers["Content-Type"] = "application/json";
      init.body = JSON.stringify(data);
    }

    try {
      const response = await fetch(url, init);

      if (!response.ok) {
        let apiError: GmooError | null = null;
        try {
          const body = await response.text();
          try { apiError = JSON.parse(body); } catch { /* not JSON */ }
        } catch {
          // Couldn't read response body
        }
        throw new GmooApiError(response.status, apiError);
      }

      return (await response.json()) as T;
    } catch (err) {
      if (err instanceof GmooApiError) throw err;
      if (signal?.aborted) throw new GmooCancelledError();
      if (controller.signal.aborted) throw new GmooTimeoutError(this.timeoutMs);
      throw err;
    } finally {
      if (timeoutId !== null) clearTimeout(timeoutId);
      signal?.removeEventListener("abort", onCallerAbort);
    }
  }

  private shouldRetry(status: number): boolean {
    return status >= 500 || status === 429;
  }

  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }
}
