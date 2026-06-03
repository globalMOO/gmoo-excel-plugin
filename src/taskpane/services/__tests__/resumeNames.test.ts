import { resolveResumeNames } from "../resumeNames";

describe("resolveResumeNames", () => {
  it("prefers same-project in-state names", () => {
    const { inputNames, outcomeNames } = resolveResumeNames({
      inputCount: 2,
      outputCount: 1,
      sameProject: true,
      stateInputNames: ["Temp", "Pressure"],
      stateOutcomeNames: ["Yield"],
      savedVariableNames: ["X1", "X2"],
      savedOutcomeNames: ["Y1"],
    });
    expect(inputNames).toEqual(["Temp", "Pressure"]);
    expect(outcomeNames).toEqual(["Yield"]);
  });

  it("recovers names from the saved state sheet when state is empty (cross-project)", () => {
    const { inputNames, outcomeNames } = resolveResumeNames({
      inputCount: 2,
      outputCount: 2,
      sameProject: false,
      stateInputNames: [],
      stateOutcomeNames: [],
      savedVariableNames: ["Width", "Height"],
      savedOutcomeNames: ["Area", "Perimeter"],
    });
    expect(inputNames).toEqual(["Width", "Height"]);
    expect(outcomeNames).toEqual(["Area", "Perimeter"]);
  });

  it("falls back to generic labels when nothing is available", () => {
    const { inputNames, outcomeNames } = resolveResumeNames({
      inputCount: 2,
      outputCount: 1,
      sameProject: false,
      stateInputNames: [],
      stateOutcomeNames: [],
    });
    expect(inputNames).toEqual(["Input 1", "Input 2"]);
    expect(outcomeNames).toEqual(["Outcome 1"]);
  });

  it("ignores blank/whitespace names and falls through to the next source", () => {
    const { inputNames } = resolveResumeNames({
      inputCount: 2,
      outputCount: 0,
      sameProject: true,
      stateInputNames: ["  ", ""],
      stateOutcomeNames: [],
      savedVariableNames: ["Real", "   "],
    });
    // index 0: blank state → saved "Real"; index 1: blank state + blank saved → generic
    expect(inputNames).toEqual(["Real", "Input 2"]);
  });
});
