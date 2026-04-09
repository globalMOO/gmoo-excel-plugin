import { useState, useCallback, useRef } from "react";
import type { GmooClient } from "../services/gmooApi";
import type { Inverse } from "../types/gmoo";
import { shouldStop, getStopReason } from "../types/gmoo";
import type { EvalConfig } from "../services/excelService";
import { evaluateCase } from "../services/excelService";

export interface OptimizationState {
  isRunning: boolean;
  iterations: Inverse[];
  currentIteration: number;
  bestL1Norm: number | null;
  initialL1Norm: number | null;
  stopReason: string | null;
  error: string | null;
}

export function useOptimization(
  client: GmooClient | null,
  objectiveId: number | null,
  evalConfig: EvalConfig | null
) {
  const [state, setState] = useState<OptimizationState>({
    isRunning: false,
    iterations: [],
    currentIteration: 0,
    bestL1Norm: null,
    initialL1Norm: null,
    stopReason: null,
    error: null,
  });

  const abortRef = useRef(false);

  const runIteration = useCallback(async (): Promise<Inverse | null> => {
    if (!client || !objectiveId || !evalConfig) {
      console.log("[Optimize] Skipping: missing", {
        client: !!client, objectiveId, evalConfig: !!evalConfig,
      });
      return null;
    }

    // 1. Suggest next inverse
    console.log(`[Optimize] suggestInverse(${objectiveId})`);
    const inverse = await client.suggestInverse(objectiveId);
    console.log(`[Optimize] Inverse id=${inverse.id}, inputs=${inverse.input.length}, loadedAt=${inverse.loadedAt}`);

    // 2. Evaluate in Excel
    const result = await evaluateCase(evalConfig, inverse.input);
    if (result.errors.length > 0) {
      throw new Error(`Formula errors: ${result.errors.join("; ")}`);
    }

    // 3. Load output back to API
    console.log(`[Optimize] loadInverseOutput(${inverse.id}, [${result.outputs.slice(0, 3).join(",")}...])`);
    const loaded = await client.loadInverseOutput(inverse.id, result.outputs);
    console.log(`[Optimize] Done: l1Norm=${loaded.l1Norm}`);

    return loaded;
  }, [client, objectiveId, evalConfig]);

  // Fetch iteration 0 from the objective's initial inverse
  const fetchInitialInverse = useCallback(async (): Promise<Inverse | null> => {
    if (!client || !objectiveId) return null;
    try {
      console.log(`[Optimize] Fetching objective ${objectiveId} for initial inverse`);
      const objective = await client.getObjective(objectiveId);
      if (objective.inverses && objective.inverses.length > 0) {
        const initial = objective.inverses[0];
        console.log(`[Optimize] Initial inverse: id=${initial.id}, l1Norm=${initial.l1Norm}`);
        return initial;
      }
    } catch (err) {
      console.warn("[Optimize] Could not fetch initial inverse:", err);
    }
    return null;
  }, [client, objectiveId]);

  const updateStateWithInverse = (inverse: Inverse, iterationNum: number, isInitial = false) => {
    setState((prev) => {
      const iterations = [...prev.iterations, inverse];
      const bestL1Norm =
        prev.bestL1Norm === null
          ? inverse.l1Norm
          : Math.min(prev.bestL1Norm, inverse.l1Norm);
      const initialL1Norm = isInitial ? inverse.l1Norm : prev.initialL1Norm;
      const stopReason = shouldStop(inverse)
        ? (() => {
            const r = getStopReason(inverse);
            return r === 1 ? "Objective satisfied"
              : r === 2 ? "Stopped (duplicate inputs)"
              : r === 3 ? "Exhausted" : "Unknown";
          })()
        : null;
      return { ...prev, iterations, currentIteration: iterationNum, bestL1Norm, initialL1Norm, stopReason };
    });
  };

  const run = useCallback(
    async (maxIterations: number) => {
      if (!client || !objectiveId || !evalConfig) return;

      abortRef.current = false;
      setState((prev) => ({
        ...prev,
        isRunning: true,
        error: null,
        stopReason: null,
      }));

      try {
        // Fetch iteration 0 if we haven't started yet
        let startFrom: number;
        const currentCount = state.iterations.length;
        if (currentCount === 0) {
          console.log(`[Optimize] run(${maxIterations}), objectiveId=${objectiveId}, fresh start`);
          const initial = await fetchInitialInverse();
          if (initial) {
            updateStateWithInverse(initial, 0, true);
          }
          startFrom = 0;
        } else {
          console.log(`[Optimize] run(${maxIterations}), objectiveId=${objectiveId}, resuming from ${currentCount}`);
          startFrom = currentCount - 1; // -1 because iteration 0 is in there
        }

        for (let i = 0; i < maxIterations; i++) {
          if (abortRef.current) {
            setState((prev) => ({ ...prev, stopReason: "Paused by user" }));
            break;
          }

          const inverse = await runIteration();
          if (!inverse) break;

          updateStateWithInverse(inverse, startFrom + i + 1);

          if (shouldStop(inverse)) break;
        }
      } catch (err) {
        console.error("[Optimize] Error:", err);
        setState((prev) => ({
          ...prev,
          error: err instanceof Error ? err.message : "Unknown error",
        }));
      } finally {
        setState((prev) => ({ ...prev, isRunning: false }));
      }
    },
    [client, objectiveId, evalConfig, runIteration, fetchInitialInverse, state.iterations.length]
  );

  const runSingle = useCallback(async () => {
    if (!client || !objectiveId || !evalConfig) return;

    setState((prev) => ({ ...prev, isRunning: true, error: null }));
    try {
      // Fetch iteration 0 if first run
      if (state.iterations.length === 0) {
        const initial = await fetchInitialInverse();
        if (initial) {
          updateStateWithInverse(initial, 0, true);
        }
      }

      const inverse = await runIteration();
      if (inverse) {
        const iterNum = state.iterations.length; // already includes iter 0 if fetched above
        updateStateWithInverse(inverse, iterNum);
      }
    } catch (err) {
      console.error("[Optimize] Single error:", err);
      setState((prev) => ({
        ...prev,
        error: err instanceof Error ? err.message : "Unknown error",
      }));
    } finally {
      setState((prev) => ({ ...prev, isRunning: false }));
    }
  }, [client, objectiveId, evalConfig, runIteration, fetchInitialInverse, state.iterations.length]);

  const stop = useCallback(() => {
    abortRef.current = true;
  }, []);

  const reset = useCallback(() => {
    abortRef.current = true;
    setState({
      isRunning: false,
      iterations: [],
      currentIteration: 0,
      bestL1Norm: null,
      initialL1Norm: null,
      stopReason: null,
      error: null,
    });
  }, []);

  return { ...state, run, stop, reset, runSingleIteration: runSingle };
}
