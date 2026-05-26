import { useState, useCallback, useEffect, useRef } from "react";
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

const EMPTY_STATE: OptimizationState = {
  isRunning: false,
  iterations: [],
  currentIteration: 0,
  bestL1Norm: null,
  initialL1Norm: null,
  stopReason: null,
  error: null,
};

function stopReasonFor(inv: Inverse): string | null {
  if (!shouldStop(inv)) return null;
  const r = getStopReason(inv);
  return r === 1 ? "Objective satisfied"
    : r === 2 ? "Stopped (duplicate inputs)"
    : r === 3 ? "Exhausted" : "Unknown";
}

export function useOptimization(
  client: GmooClient | null,
  objectiveId: number | null,
  evalConfig: EvalConfig | null
) {
  const [state, setState] = useState<OptimizationState>(EMPTY_STATE);

  const abortRef = useRef(false);
  // Track which objectiveId we've already loaded history for, so a re-render
  // (e.g. evalConfig flipping in) doesn't re-fetch and clobber an in-progress
  // run.
  const loadedHistoryForRef = useRef<number | null>(null);

  // When objectiveId changes, the iterations from the prior objective no
  // longer apply. Reset, then asynchronously load any existing inverses from
  // the API so reopening a workbook or switching objectives shows the chart
  // immediately rather than blank.
  useEffect(() => {
    abortRef.current = true; // cancel any running loop tied to the prior id
    setState(EMPTY_STATE);
    if (!client || !objectiveId) {
      loadedHistoryForRef.current = null;
      return;
    }
    if (loadedHistoryForRef.current === objectiveId) return;
    loadedHistoryForRef.current = objectiveId;
    let cancelled = false;
    client
      .getObjective(objectiveId)
      .then((obj) => {
        if (cancelled || !obj.inverses || obj.inverses.length === 0) return;
        const sorted = [...obj.inverses].sort((a, b) => a.iteration - b.iteration);
        const last = sorted[sorted.length - 1];
        setState({
          isRunning: false,
          iterations: sorted,
          currentIteration: last.iteration,
          bestL1Norm: Math.min(...sorted.map((i) => i.l1Norm)),
          initialL1Norm: sorted[0].l1Norm,
          stopReason: stopReasonFor(last),
          error: null,
        });
      })
      .catch((err: unknown) => {
        if (cancelled) return;
        // Non-fatal — user can still run new iterations from scratch.
        console.warn("[Optimize] Could not load history:", err);
      });
    return () => {
      cancelled = true;
    };
  }, [client, objectiveId]);

  const runIteration = useCallback(async (): Promise<Inverse | null> => {
    if (!client || !objectiveId || !evalConfig) return null;

    const inverse = await client.suggestInverse(objectiveId);

    const result = await evaluateCase(evalConfig, inverse.input);
    if (result.errors.length > 0) {
      throw new Error(`Formula errors: ${result.errors.join("; ")}`);
    }

    return client.loadInverseOutput(inverse.id, result.outputs);
  }, [client, objectiveId, evalConfig]);

  // Fetch iteration 0 from the objective's initial inverse
  const fetchInitialInverse = useCallback(async (): Promise<Inverse | null> => {
    if (!client || !objectiveId) return null;
    try {
      const objective = await client.getObjective(objectiveId);
      if (objective.inverses && objective.inverses.length > 0) {
        return objective.inverses[0];
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
      return {
        ...prev,
        iterations,
        currentIteration: iterationNum,
        bestL1Norm,
        initialL1Norm,
        stopReason: stopReasonFor(inverse),
      };
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
          const initial = await fetchInitialInverse();
          if (initial) {
            updateStateWithInverse(initial, 0, true);
          }
          startFrom = 0;
        } else {
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
      // Track count locally — `state.iterations.length` is captured in this
      // closure and won't reflect the in-flight `updateStateWithInverse` call
      // above. Reading it post-setState would underreport by 1 on the very
      // first manual iteration, mis-labeling the new inverse as iter 0.
      let count = state.iterations.length;
      if (count === 0) {
        const initial = await fetchInitialInverse();
        if (initial) {
          updateStateWithInverse(initial, 0, true);
          count = 1;
        }
      }

      const inverse = await runIteration();
      if (inverse) {
        updateStateWithInverse(inverse, count);
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
    // Allow the next render to re-load history for the current objective
    // (e.g. user came back to Optimize after editing the objective).
    loadedHistoryForRef.current = null;
    setState(EMPTY_STATE);
  }, []);

  return { ...state, run, stop, reset, runSingleIteration: runSingle };
}
