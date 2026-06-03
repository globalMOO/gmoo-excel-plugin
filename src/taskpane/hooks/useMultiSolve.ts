// Multi-Solve: runs N independent optimizations with randomized initial inputs
// and reports every converged solution. No dedupe — a normalized-Euclidean
// "uniqueness threshold" is arbitrary in input-space (depends on dimensionality,
// ignores output behavior), so we let the user see all of them and judge.
// The "Dist. to nearest" column in the sheet log is retained as a purely
// diagnostic field to show how clustered the results are.
import { useCallback, useRef, useState } from "react";
import type { GmooClient } from "../services/gmooApi";
import type { Inverse, Result } from "../types/gmoo";
import {
  getStopReason,
  shouldStop,
  isSolvedStop,
} from "../types/gmoo";
import type { EvalConfig } from "../services/excelService";
import { evaluateCase, writeMultiSolveRun } from "../services/excelService";
import { randomInputWithinBounds } from "../services/sampling";
import type { InputVariable } from "../types/workbookState";

export interface MultiSolveSolution {
  runIndex: number;       // which random-start produced it (0..numRuns-1)
  input: number[];
  output: number[];
  results: Result[];      // per-outcome breakdown from the best inverse
  l1Norm: number;         // filtered L1 norm for the best inverse
  satisfied: boolean;
  iterations: number;     // iteration index of the best inverse
}

export interface MultiSolveConfig {
  targets: number[];
  types: string[];
  minBounds: number[];
  maxBounds: number[];
  numRuns?: number;         // default 10
  maxIterations?: number;   // default 50
}

export interface MultiSolveProgress {
  run: number;             // 1-based
  totalRuns: number;
  stage: "randomizing" | "iterating" | "done";
  iteration?: number;
}

export interface MultiSolveState {
  isRunning: boolean;
  solutions: MultiSolveSolution[];
  runsCompleted: number;     // how many runs produced a final inverse (successes)
  runsFailed: number;        // how many runs errored out and produced nothing
  progress: MultiSolveProgress | null;
  error: string | null;
}

const DEFAULT_NUM_RUNS = 10;
const DEFAULT_MAX_ITERATIONS = 50;

function normalizeInputs(input: number[], inputVars: InputVariable[]): number[] {
  return input.map((v, i) => {
    const range = inputVars[i].max - inputVars[i].min;
    if (range === 0) return 0;
    return (v - inputVars[i].min) / range;
  });
}

function euclidean(a: number[], b: number[]): number {
  let s = 0;
  for (let i = 0; i < a.length; i++) {
    const d = a[i] - b[i];
    s += d * d;
  }
  return Math.sqrt(s);
}

export function useMultiSolve(
  client: GmooClient | null,
  trialId: number | null,
  evalConfig: EvalConfig | null,
  inputVariables: InputVariable[],
  outcomeNames: string[]
) {
  const [state, setState] = useState<MultiSolveState>({
    isRunning: false,
    solutions: [],
    runsCompleted: 0,
    runsFailed: 0,
    progress: null,
    error: null,
  });

  const abortRef = useRef(false);

  const stop = useCallback(() => {
    abortRef.current = true;
  }, []);

  const reset = useCallback(() => {
    abortRef.current = true;
    setState({
      isRunning: false,
      solutions: [],
      runsCompleted: 0,
      runsFailed: 0,
      progress: null,
      error: null,
    });
  }, []);

  const run = useCallback(
    async (config: MultiSolveConfig) => {
      if (!client || !trialId || !evalConfig || inputVariables.length === 0) {
        setState((prev) => ({
          ...prev,
          error: "Missing client, trial, evaluation config, or input variables.",
        }));
        return;
      }

      const numRuns = config.numRuns ?? DEFAULT_NUM_RUNS;
      const maxIterations = config.maxIterations ?? DEFAULT_MAX_ITERATIONS;

      abortRef.current = false;
      setState({
        isRunning: true,
        solutions: [],
        runsCompleted: 0,
        runsFailed: 0,
        progress: { run: 0, totalRuns: numRuns, stage: "randomizing" },
        error: null,
      });

      const collected: MultiSolveSolution[] = [];
      let runsCompleted = 0;
      let runsFailed = 0;
      let sheetInitialized = false;
      const inputVarNames = inputVariables.map((v) => v.name);

      try {
        // Write a single diagnostic row per run to the "GMOO Multi-Solve" sheet.
        // Wrapped so we can centralize the sheet-reset flag and error handling.
        const logRun = async (
          runIndex: number,
          status: "unique" | "failed",
          solutionIdx: number | null,
          distanceToNearest: number | null,
          satisfied: boolean,
          l1Norm: number | null,
          iterations: number | null,
          initialInputs: number[],
          finalInputs: number[] | null,
          outputs: number[] | null,
          note?: string
        ) => {
          try {
            await writeMultiSolveRun(
              {
                runIndex,
                status,
                solutionIdx,
                distanceToNearest,
                satisfied,
                l1Norm,
                iterations,
                initialInputs,
                finalInputs,
                outputs,
                note,
              },
              inputVarNames,
              outcomeNames,
              !sheetInitialized
            );
            sheetInitialized = true;
          } catch (err) {
            console.warn("[MultiSolve] writeMultiSolveRun failed", err);
          }
        };

        for (let run = 0; run < numRuns; run++) {
          if (abortRef.current) break;

          setState((prev) => ({
            ...prev,
            progress: { run: run + 1, totalRuns: numRuns, stage: "randomizing" },
          }));

          // 1. Random initial input within bounds
          const initialInput = randomInputWithinBounds(inputVariables);

          // 2. Evaluate initial input through Excel
          let initialOutput: number[];
          try {
            const evalResult = await evaluateCase(evalConfig, initialInput);
            if (evalResult.errors.length > 0) {
              console.warn(`[MultiSolve] Run ${run + 1}: initial eval errors`, evalResult.errors);
              runsFailed++;
              await logRun(run, "failed", null, null, false, null, null, initialInput, null, null, `initial eval errors: ${evalResult.errors.join("; ")}`);
              setState((prev) => ({ ...prev, runsFailed }));
              continue;
            }
            initialOutput = evalResult.outputs;
          } catch (err) {
            console.warn(`[MultiSolve] Run ${run + 1}: initial eval failed`, err);
            runsFailed++;
            await logRun(run, "failed", null, null, false, null, null, initialInput, null, null, `initial eval threw: ${err instanceof Error ? err.message : String(err)}`);
            setState((prev) => ({ ...prev, runsFailed }));
            continue;
          }

          if (abortRef.current) break;

          // 3. Create a fresh objective for this run
          let objective;
          try {
            objective = await client.loadObjectives(
              trialId,
              config.targets,
              config.types,
              initialInput,
              initialOutput,
              0,
              config.minBounds,
              config.maxBounds
            );
          } catch (err) {
            console.warn(`[MultiSolve] Run ${run + 1}: loadObjectives failed`, err);
            runsFailed++;
            await logRun(run, "failed", null, null, false, null, null, initialInput, null, initialOutput, `loadObjectives threw: ${err instanceof Error ? err.message : String(err)}`);
            setState((prev) => ({ ...prev, runsFailed }));
            continue;
          }

          // 4. Iterate — track the best inverse (lowest filtered L1 norm)
          let bestInverse: Inverse | null =
            objective.inverses && objective.inverses.length > 0
              ? objective.inverses[0]
              : null;
          let lastInverse: Inverse | null = bestInverse;

          for (let i = 0; i < maxIterations; i++) {
            if (abortRef.current) break;

            setState((prev) => ({
              ...prev,
              progress: {
                run: run + 1,
                totalRuns: numRuns,
                stage: "iterating",
                iteration: i + 1,
              },
            }));

            let suggested: Inverse;
            try {
              suggested = await client.suggestInverse(objective.id);
            } catch (err) {
              console.warn(`[MultiSolve] Run ${run + 1} iter ${i + 1}: suggestInverse failed`, err);
              break;
            }

            let excelResult;
            try {
              excelResult = await evaluateCase(evalConfig, suggested.input);
            } catch (err) {
              console.warn(`[MultiSolve] Run ${run + 1} iter ${i + 1}: eval failed`, err);
              break;
            }
            if (excelResult.errors.length > 0) {
              console.warn(`[MultiSolve] Run ${run + 1} iter ${i + 1}: eval errors`, excelResult.errors);
              break;
            }

            try {
              lastInverse = await client.loadInverseOutput(suggested.id, excelResult.outputs);
            } catch (err) {
              console.warn(`[MultiSolve] Run ${run + 1} iter ${i + 1}: loadInverseOutput failed`, err);
              break;
            }

            // Rank by the API's l1Norm (the score the optimizer is actually
            // minimizing). filteredL1Norm degenerates to 0 when all objectives
            // are inequality/minimize/maximize, which would pin `bestInverse`
            // to the random initial point and report it as a "solution".
            if (
              bestInverse === null ||
              lastInverse.l1Norm < bestInverse.l1Norm
            ) {
              bestInverse = lastInverse;
            }

            if (shouldStop(lastInverse)) break;
          }

          if (abortRef.current) break;

          // 5. Record the run. Every converged inverse is reported — no dedupe.
          if (!bestInverse) {
            runsFailed++;
            await logRun(run, "failed", null, null, false, null, null, initialInput, null, null, "no inverse produced");
            setState((prev) => ({ ...prev, runsFailed }));
            continue;
          }

          runsCompleted++;

          // Nearest-neighbor distance in input space — kept as a diagnostic
          // column so the user can see how clustered the results are.
          const normalized = normalizeInputs(bestInverse.input, inputVariables);
          let nearestDist: number | null = null;
          for (let s = 0; s < collected.length; s++) {
            const normSol = normalizeInputs(collected[s].input, inputVariables);
            const d = euclidean(normalized, normSol);
            if (nearestDist === null || d < nearestDist) {
              nearestDist = d;
            }
          }

          const satisfied =
            lastInverse !== null && isSolvedStop(getStopReason(lastInverse));

          const newSolution: MultiSolveSolution = {
            runIndex: run,
            input: bestInverse.input,
            output: bestInverse.output,
            results: bestInverse.results,
            l1Norm: bestInverse.l1Norm,
            satisfied,
            iterations: bestInverse.iteration,
          };
          collected.push(newSolution);

          await logRun(
            run,
            "unique",
            collected.length,
            nearestDist,
            satisfied,
            bestInverse.l1Norm,
            bestInverse.iteration,
            initialInput,
            bestInverse.input,
            bestInverse.output
          );

          setState((prev) => ({
            ...prev,
            solutions: [...collected],
            runsCompleted,
          }));
        }

        setState((prev) => ({
          ...prev,
          isRunning: false,
          solutions: [...collected],
          runsCompleted,
          runsFailed,
          progress: {
            run: Math.min(numRuns, prev.progress?.run ?? numRuns),
            totalRuns: numRuns,
            stage: "done",
          },
        }));
      } catch (err) {
        console.error("[MultiSolve] Unexpected error", err);
        setState((prev) => ({
          ...prev,
          isRunning: false,
          error: err instanceof Error ? err.message : "Multi-solve failed.",
        }));
      }
    },
    [client, trialId, evalConfig, inputVariables, outcomeNames]
  );

  return { ...state, run, stop, reset };
}
