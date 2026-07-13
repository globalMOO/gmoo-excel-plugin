import { renderHook, act, waitFor } from "@testing-library/react";
import { useOptimization } from "../useOptimization";
import type { GmooClient } from "../../services/gmooApi";
import { GmooCancelledError } from "../../services/gmooApi";
import type { EvalConfig } from "../../services/excelService";
import { evaluateCase } from "../../services/excelService";
import type { Inverse, Objective } from "../../types/gmoo";

// excelService touches the Excel global at call time; replace it so the hook
// can run under jsdom.
jest.mock("../../services/excelService", () => ({
  evaluateCase: jest.fn(),
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

// Only `inverses` is read off the objective by the hook.
function makeObjective(inverses: Inverse[]): Objective {
  return { inverses } as unknown as Objective;
}

function makeClient() {
  return {
    getObjective: jest.fn(),
    suggestInverse: jest.fn(),
    loadInverseOutput: jest.fn(),
  };
}

type MockClient = ReturnType<typeof makeClient>;
const asClient = (c: MockClient) => c as unknown as GmooClient;

const evalConfig = { variableCount: 2, outcomeCount: 1 } as EvalConfig;

beforeEach(() => {
  mockEvaluateCase.mockReset();
  mockEvaluateCase.mockResolvedValue({ outputs: [42], errors: [] });
  jest.spyOn(console, "warn").mockImplementation(() => {});
  jest.spyOn(console, "error").mockImplementation(() => {});
});

afterEach(() => {
  jest.restoreAllMocks();
});

describe("history loading", () => {
  it("loads and sorts existing inverses when an objective id is set", async () => {
    const client = makeClient();
    client.getObjective.mockResolvedValue(
      makeObjective([
        makeInverse({ iteration: 2, l1Norm: 0.5 }),
        makeInverse({ iteration: 0, l1Norm: 2 }),
        makeInverse({ iteration: 1, l1Norm: 1 }),
      ])
    );

    const { result } = renderHook(() => useOptimization(asClient(client), 7, evalConfig));

    await waitFor(() => expect(result.current.iterations).toHaveLength(3));
    expect(result.current.iterations.map((i) => i.iteration)).toEqual([0, 1, 2]);
    expect(result.current.currentIteration).toBe(2);
    expect(result.current.bestL1Norm).toBe(0.5);
    expect(result.current.initialL1Norm).toBe(2);
    expect(result.current.stopReason).toBeNull();
    expect(result.current.isRunning).toBe(false);
  });

  it("maps a converged last inverse to the converged stop reason", async () => {
    const client = makeClient();
    client.getObjective.mockResolvedValue(
      makeObjective([
        makeInverse({ iteration: 0, l1Norm: 2 }),
        makeInverse({ iteration: 1, l1Norm: 1, stoppedAt: "2026-01-02T00:00:00Z" }),
      ])
    );

    const { result } = renderHook(() => useOptimization(asClient(client), 7, evalConfig));

    await waitFor(() =>
      expect(result.current.stopReason).toBe(
        "Converged (optimizer reached the best achievable solution)"
      )
    );
  });

  it("treats a failed history load as non-fatal", async () => {
    const client = makeClient();
    client.getObjective.mockRejectedValue(new Error("network down"));

    const { result } = renderHook(() => useOptimization(asClient(client), 7, evalConfig));

    await waitFor(() => expect(console.warn).toHaveBeenCalled());
    expect(result.current.iterations).toHaveLength(0);
    expect(result.current.error).toBeNull();
    expect(result.current.isRunning).toBe(false);
  });

  it("does not re-fetch history when unrelated props change, but does for a new objective", async () => {
    const client = makeClient();
    client.getObjective.mockResolvedValue(makeObjective([makeInverse({ iteration: 0 })]));

    const { result, rerender } = renderHook(
      ({ objectiveId, config }) => useOptimization(asClient(client), objectiveId, config),
      { initialProps: { objectiveId: 1, config: null as EvalConfig | null } }
    );
    await waitFor(() => expect(result.current.iterations).toHaveLength(1));
    expect(client.getObjective).toHaveBeenCalledTimes(1);

    // evalConfig flipping in (same objective) must not re-fetch and clobber
    rerender({ objectiveId: 1, config: evalConfig });
    await act(async () => {});
    expect(client.getObjective).toHaveBeenCalledTimes(1);

    // a different objective resets state and loads its own history
    client.getObjective.mockResolvedValue(
      makeObjective([makeInverse({ iteration: 0 }), makeInverse({ iteration: 1 })])
    );
    rerender({ objectiveId: 2, config: evalConfig });
    await waitFor(() => expect(result.current.iterations).toHaveLength(2));
    expect(client.getObjective).toHaveBeenCalledTimes(2);
  });
});

describe("run()", () => {
  function setupFreshRun(client: MockClient) {
    // First getObjective call is the mount history load (empty); later calls
    // serve fetchInitialInverse with iteration 0.
    client.getObjective
      .mockResolvedValueOnce(makeObjective([]))
      .mockResolvedValue(makeObjective([makeInverse({ iteration: 0, l1Norm: 10 })]));
  }

  it("fetches iteration 0, then iterates until the API reports satisfied", async () => {
    const client = makeClient();
    setupFreshRun(client);
    client.suggestInverse.mockResolvedValue(makeInverse({ input: [5, 6] }));
    client.loadInverseOutput
      .mockResolvedValueOnce(makeInverse({ iteration: 1, l1Norm: 5 }))
      .mockResolvedValueOnce(
        makeInverse({ iteration: 2, l1Norm: 0, satisfiedAt: "2026-01-02T00:00:00Z" })
      );

    const { result } = renderHook(() => useOptimization(asClient(client), 7, evalConfig));
    await act(async () => {}); // let the (empty) history load settle

    await act(async () => {
      await result.current.run(10);
    });

    expect(result.current.iterations).toHaveLength(3); // initial + 2 iterations
    expect(result.current.currentIteration).toBe(2);
    expect(result.current.stopReason).toBe("Objective satisfied");
    expect(result.current.isRunning).toBe(false);
    expect(result.current.error).toBeNull();
    expect(result.current.initialL1Norm).toBe(10);
    expect(result.current.bestL1Norm).toBe(0);
    expect(client.suggestInverse).toHaveBeenCalledTimes(2);
    expect(mockEvaluateCase).toHaveBeenCalledWith(evalConfig, [5, 6]);
  });

  it("stops at maxIterations when the objective is not satisfied", async () => {
    const client = makeClient();
    setupFreshRun(client);
    client.suggestInverse.mockResolvedValue(makeInverse());
    let iter = 0;
    client.loadInverseOutput.mockImplementation(async () =>
      makeInverse({ iteration: ++iter, l1Norm: 5 })
    );

    const { result } = renderHook(() => useOptimization(asClient(client), 7, evalConfig));
    await act(async () => {});

    await act(async () => {
      await result.current.run(3);
    });

    expect(client.suggestInverse).toHaveBeenCalledTimes(3);
    expect(result.current.iterations).toHaveLength(4); // initial + 3
    expect(result.current.isRunning).toBe(false);
    expect(result.current.stopReason).toBeNull();
  });

  it("surfaces formula errors and always clears isRunning", async () => {
    const client = makeClient();
    setupFreshRun(client);
    client.suggestInverse.mockResolvedValue(makeInverse());
    mockEvaluateCase.mockResolvedValue({
      outputs: [0],
      errors: ["Outcome 1 (Sheet1!B2): #DIV/0!"],
    });

    const { result } = renderHook(() => useOptimization(asClient(client), 7, evalConfig));
    await act(async () => {});

    await act(async () => {
      await result.current.run(5);
    });

    expect(result.current.error).toBe("Formula errors: Outcome 1 (Sheet1!B2): #DIV/0!");
    expect(result.current.isRunning).toBe(false);
    expect(client.loadInverseOutput).not.toHaveBeenCalled();
  });

  it("surfaces API failures and always clears isRunning", async () => {
    const client = makeClient();
    setupFreshRun(client);
    client.suggestInverse.mockRejectedValue(new Error("API error 500: Unknown error"));

    const { result } = renderHook(() => useOptimization(asClient(client), 7, evalConfig));
    await act(async () => {});

    await act(async () => {
      await result.current.run(5);
    });

    expect(result.current.error).toBe("API error 500: Unknown error");
    expect(result.current.isRunning).toBe(false);
  });

  it("stop() halts the loop with 'Paused by user' and runs no further iterations", async () => {
    const client = makeClient();
    setupFreshRun(client);
    client.suggestInverse.mockResolvedValue(makeInverse());

    const { result } = renderHook(() => useOptimization(asClient(client), 7, evalConfig));
    await act(async () => {});

    client.loadInverseOutput.mockImplementation(async () => {
      result.current.stop(); // user clicks Stop while iteration 1 finishes
      return makeInverse({ iteration: 1, l1Norm: 5 });
    });

    await act(async () => {
      await result.current.run(10);
    });

    expect(client.suggestInverse).toHaveBeenCalledTimes(1);
    expect(result.current.iterations).toHaveLength(2); // initial + the finished iteration
    expect(result.current.stopReason).toBe("Paused by user");
    expect(result.current.isRunning).toBe(false);
  });

  it("stop() aborts an in-flight API call and reports 'Paused by user', not an error", async () => {
    const client = makeClient();
    setupFreshRun(client);
    // Mirrors the real client: the call hangs until the run's signal aborts,
    // then rejects with GmooCancelledError.
    client.suggestInverse.mockImplementation(
      (_id: number, signal?: AbortSignal) =>
        new Promise((_, reject) =>
          signal?.addEventListener("abort", () => reject(new GmooCancelledError()))
        )
    );

    const { result } = renderHook(() => useOptimization(asClient(client), 7, evalConfig));
    await act(async () => {});

    await act(async () => {
      const running = result.current.run(10);
      await Promise.resolve(); // let the loop reach the hanging suggestInverse
      result.current.stop();
      await running;
    });

    expect(result.current.stopReason).toBe("Paused by user");
    expect(result.current.error).toBeNull();
    expect(result.current.isRunning).toBe(false);
  });

  it("resumes iteration numbering from existing history", async () => {
    const client = makeClient();
    client.getObjective.mockResolvedValue(
      makeObjective([
        makeInverse({ iteration: 0, l1Norm: 10 }),
        makeInverse({ iteration: 1, l1Norm: 5 }),
      ])
    );
    client.suggestInverse.mockResolvedValue(makeInverse());
    client.loadInverseOutput.mockResolvedValue(
      makeInverse({ iteration: 2, l1Norm: 1, satisfiedAt: "2026-01-02T00:00:00Z" })
    );

    const { result } = renderHook(() => useOptimization(asClient(client), 7, evalConfig));
    await waitFor(() => expect(result.current.iterations).toHaveLength(2));

    await act(async () => {
      await result.current.run(10);
    });

    // history had iterations 0..1, so the next one is labeled 2 — no re-fetch
    // of iteration 0 and no duplicate entries
    expect(result.current.iterations).toHaveLength(3);
    expect(result.current.currentIteration).toBe(2);
    expect(client.getObjective).toHaveBeenCalledTimes(1); // history load only
  });
});

describe("runSingleIteration()", () => {
  it("fetches iteration 0 first and labels the manual iteration 1", async () => {
    const client = makeClient();
    client.getObjective
      .mockResolvedValueOnce(makeObjective([]))
      .mockResolvedValue(makeObjective([makeInverse({ iteration: 0, l1Norm: 10 })]));
    client.suggestInverse.mockResolvedValue(makeInverse());
    client.loadInverseOutput.mockResolvedValue(makeInverse({ iteration: 1, l1Norm: 5 }));

    const { result } = renderHook(() => useOptimization(asClient(client), 7, evalConfig));
    await act(async () => {});

    await act(async () => {
      await result.current.runSingleIteration();
    });

    expect(result.current.iterations).toHaveLength(2);
    expect(result.current.currentIteration).toBe(1);
    expect(result.current.isRunning).toBe(false);
  });
});

describe("reset()", () => {
  it("clears state and allows history to reload for the same objective", async () => {
    const client = makeClient();
    client.getObjective.mockResolvedValue(makeObjective([makeInverse({ iteration: 0 })]));

    const { result, rerender } = renderHook(
      ({ c }) => useOptimization(asClient(c), 7, evalConfig),
      { initialProps: { c: client } }
    );
    await waitFor(() => expect(result.current.iterations).toHaveLength(1));

    act(() => {
      result.current.reset();
    });
    expect(result.current.iterations).toHaveLength(0);

    // The history effect re-runs when the client identity changes. Without
    // reset() it would skip the fetch (same objective already loaded); after
    // reset() it must fetch again.
    const client2 = makeClient();
    client2.getObjective.mockResolvedValue(makeObjective([makeInverse({ iteration: 0 })]));
    rerender({ c: client2 });
    await waitFor(() => expect(result.current.iterations).toHaveLength(1));
    expect(client2.getObjective).toHaveBeenCalledTimes(1);
  });
});
