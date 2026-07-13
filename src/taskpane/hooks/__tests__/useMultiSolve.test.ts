import { renderHook, act } from "@testing-library/react";
import { useMultiSolve, MultiSolveConfig } from "../useMultiSolve";
import type { GmooClient } from "../../services/gmooApi";
import { GmooCancelledError } from "../../services/gmooApi";
import type { EvalConfig } from "../../services/excelService";
import { evaluateCase } from "../../services/excelService";
import type { Inverse, Objective } from "../../types/gmoo";
import type { InputVariable } from "../../types/workbookState";

// excelService touches the Excel global at call time; replace it so the hook
// can run under jsdom.
jest.mock("../../services/excelService", () => ({
  evaluateCase: jest.fn(),
  writeMultiSolveRun: jest.fn(),
}));
jest.mock("../../services/sampling", () => ({
  randomInputWithinBounds: jest.fn(() => [0.5, 0.5]),
}));

const mockEvaluateCase = evaluateCase as jest.Mock;

let nextId = 1;

function makeInverse(overrides: Partial<Inverse> = {}): Inverse {
  return {
    id: nextId++,
    createdAt: "2026-01-01T00:00:00Z",
    updatedAt: "2026-01-01T00:00:00Z",
    disabledAt: null,
    loadedAt: null,
    satisfiedAt: null,
    stoppedAt: null,
    exhaustedAt: null,
    iteration: 0,
    l1Norm: 1,
    suggestTime: 0,
    computeTime: 0,
    input: [1, 2],
    output: [3],
    results: [],
    ...overrides,
  };
}

// Only `id` and `inverses` are read off the objective by the hook.
function makeObjective(inverses: Inverse[]): Objective {
  return { id: nextId++, inverses } as unknown as Objective;
}

function makeClient() {
  return {
    loadObjectives: jest.fn(),
    suggestInverse: jest.fn(),
    loadInverseOutput: jest.fn(),
  };
}

type MockClient = ReturnType<typeof makeClient>;
const asClient = (c: MockClient) => c as unknown as GmooClient;

const evalConfig = { variableCount: 2, outcomeCount: 1 } as EvalConfig;
const inputVariables: InputVariable[] = [
  { name: "x", type: "float", min: 0, max: 1 },
  { name: "y", type: "float", min: 0, max: 1 },
];
const config: MultiSolveConfig = {
  targets: [1],
  types: ["exact"],
  minBounds: [0],
  maxBounds: [0],
  numRuns: 2,
  maxIterations: 5,
};

function render(client: MockClient) {
  return renderHook(() =>
    useMultiSolve(asClient(client), 3, evalConfig, inputVariables, ["Outcome 1"])
  );
}

beforeEach(() => {
  mockEvaluateCase.mockReset();
  mockEvaluateCase.mockResolvedValue({ outputs: [42], errors: [] });
  jest.spyOn(console, "warn").mockImplementation(() => {});
  jest.spyOn(console, "error").mockImplementation(() => {});
});

afterEach(() => {
  jest.restoreAllMocks();
});

it("collects a solution per run and reports completion", async () => {
  const client = makeClient();
  client.loadObjectives.mockImplementation(async () =>
    makeObjective([makeInverse({ l1Norm: 10 })])
  );
  client.suggestInverse.mockResolvedValue(makeInverse());
  client.loadInverseOutput.mockImplementation(async () =>
    makeInverse({ iteration: 1, l1Norm: 0, satisfiedAt: "2026-01-02T00:00:00Z" })
  );

  const { result } = render(client);

  await act(async () => {
    await result.current.run(config);
  });

  expect(result.current.solutions).toHaveLength(2);
  expect(result.current.runsCompleted).toBe(2);
  expect(result.current.runsFailed).toBe(0);
  expect(result.current.solutions[0].satisfied).toBe(true);
  expect(result.current.isRunning).toBe(false);
  expect(result.current.progress?.stage).toBe("done");
});

it("counts a run that errors out as failed and moves on to the next run", async () => {
  const client = makeClient();
  client.loadObjectives
    .mockRejectedValueOnce(new Error("API error 422"))
    .mockResolvedValue(makeObjective([makeInverse({ l1Norm: 10 })]));
  client.suggestInverse.mockResolvedValue(makeInverse());
  client.loadInverseOutput.mockResolvedValue(
    makeInverse({ iteration: 1, l1Norm: 0, satisfiedAt: "2026-01-02T00:00:00Z" })
  );

  const { result } = render(client);

  await act(async () => {
    await result.current.run(config);
  });

  expect(result.current.runsFailed).toBe(1);
  expect(result.current.runsCompleted).toBe(1);
  expect(result.current.solutions).toHaveLength(1);
  expect(result.current.error).toBeNull();
});

it("stop() aborts an in-flight API call without counting it as a failed run", async () => {
  const client = makeClient();
  // Mirrors the real client: hangs until the run's signal aborts, then
  // rejects with GmooCancelledError.
  client.loadObjectives.mockImplementation(
    (...args: unknown[]) =>
      new Promise((_, reject) => {
        const signal = args[args.length - 1] as AbortSignal;
        signal.addEventListener("abort", () => reject(new GmooCancelledError()));
      })
  );

  const { result } = render(client);

  await act(async () => {
    const running = result.current.run({ ...config, numRuns: 5 });
    await Promise.resolve(); // let the first run reach the hanging call
    result.current.stop();
    await running;
  });

  expect(result.current.runsFailed).toBe(0);
  expect(result.current.runsCompleted).toBe(0);
  expect(result.current.solutions).toHaveLength(0);
  expect(result.current.error).toBeNull();
  expect(result.current.isRunning).toBe(false);
  expect(client.loadObjectives).toHaveBeenCalledTimes(1); // no further runs started
});
