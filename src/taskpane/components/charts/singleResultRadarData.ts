// Pure data builder for the single-result radar (separated from the Chart.js
// component so it can be unit-tested without a DOM).
import type { Inverse } from "../../types/gmoo";
import { isTargetBasedType } from "../../types/gmoo";

export interface SingleResultRadarData {
  inputLabels: string[];
  inputValues: number[];
  outputLabels: string[];
  achieved: number[];
  /** Target per outcome; null where the objective type has no target. */
  target: (number | null)[];
  /** True when at least one outcome has a meaningful target to overlay. */
  hasAnyTarget: boolean;
}

export function buildSingleResultRadarData(
  inverse: Inverse,
  inputVariableNames: string[],
  outcomeNames: string[]
): SingleResultRadarData {
  const inputLabels = inputVariableNames.map((n, i) => n || `Input ${i + 1}`);
  const inputValues = inputVariableNames.map((_, i) => inverse.input?.[i] ?? 0);

  const outputLabels = outcomeNames.map((n, i) => n || `Outcome ${i + 1}`);
  const achieved: number[] = [];
  const target: (number | null)[] = [];
  let hasAnyTarget = false;

  for (let i = 0; i < outcomeNames.length; i++) {
    const res = inverse.results?.[i];
    achieved.push(res?.output ?? inverse.output?.[i] ?? 0);
    if (res && isTargetBasedType(res.objectiveType)) {
      target.push(res.objective);
      hasAnyTarget = true;
    } else {
      target.push(null);
    }
  }

  return { inputLabels, inputValues, outputLabels, achieved, target, hasAnyTarget };
}
