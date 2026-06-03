import { buildSingleResultRadarData } from "../singleResultRadarData";
import type { Inverse, Result } from "../../../types/gmoo";
import { ObjectiveType } from "../../../types/gmoo";

function makeResult(partial: Partial<Result>): Result {
  return {
    id: 0,
    createdAt: "",
    updatedAt: "",
    disabledAt: null,
    number: 0,
    objective: 0,
    objectiveType: ObjectiveType.Value,
    minimumBound: 0,
    maximumBound: 0,
    output: 0,
    error: 0,
    detail: "",
    satisfied: false,
    ...partial,
  };
}

function makeInverse(partial: Partial<Inverse>): Inverse {
  return {
    id: 0,
    createdAt: "",
    updatedAt: "",
    disabledAt: null,
    loadedAt: null,
    satisfiedAt: null,
    stoppedAt: null,
    exhaustedAt: null,
    iteration: 0,
    l1Norm: 0,
    suggestTime: 0,
    computeTime: 0,
    input: [],
    output: [],
    results: [],
    ...partial,
  };
}

describe("buildSingleResultRadarData", () => {
  it("uses raw input values and falls back to generic labels", () => {
    const inv = makeInverse({ input: [1.5, 2.5, 3.5] });
    const data = buildSingleResultRadarData(inv, ["A", "", "C"], []);
    expect(data.inputLabels).toEqual(["A", "Input 2", "C"]);
    expect(data.inputValues).toEqual([1.5, 2.5, 3.5]);
  });

  it("overlays achieved vs target for target-based objectives", () => {
    const inv = makeInverse({
      output: [10, 20],
      results: [
        makeResult({ objectiveType: ObjectiveType.Value, objective: 9, output: 10 }),
        makeResult({ objectiveType: ObjectiveType.Percent, objective: 21, output: 20 }),
      ],
    });
    const data = buildSingleResultRadarData(inv, [], ["Out1", "Out2"]);
    expect(data.achieved).toEqual([10, 20]);
    expect(data.target).toEqual([9, 21]);
    expect(data.hasAnyTarget).toBe(true);
  });

  it("suppresses the target for Minimize/Maximize outcomes", () => {
    const inv = makeInverse({
      output: [5, 7],
      results: [
        makeResult({ objectiveType: ObjectiveType.Minimize, objective: 0, output: 5 }),
        makeResult({ objectiveType: ObjectiveType.Maximize, objective: 0, output: 7 }),
      ],
    });
    const data = buildSingleResultRadarData(inv, [], ["Cost", "Throughput"]);
    expect(data.achieved).toEqual([5, 7]);
    expect(data.target).toEqual([null, null]);
    expect(data.hasAnyTarget).toBe(false);
  });

  it("falls back to inverse.output when a result row is missing", () => {
    const inv = makeInverse({ output: [42], results: [] });
    const data = buildSingleResultRadarData(inv, [], ["Only"]);
    expect(data.achieved).toEqual([42]);
    expect(data.target).toEqual([null]);
  });
});
