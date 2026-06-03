import { StopReason, isSolvedStop, getStopReasonLabel } from "../gmoo";

describe("isSolvedStop", () => {
  it("treats Satisfied and Stopped (converged) as solved", () => {
    expect(isSolvedStop(StopReason.Satisfied)).toBe(true);
    expect(isSolvedStop(StopReason.Stopped)).toBe(true);
  });

  it("treats Exhausted and Running as not solved", () => {
    expect(isSolvedStop(StopReason.Exhausted)).toBe(false);
    expect(isSolvedStop(StopReason.Running)).toBe(false);
  });
});

describe("getStopReasonLabel", () => {
  it("describes Stopped as a converged/solved state, not a failure", () => {
    const label = getStopReasonLabel(StopReason.Stopped).toLowerCase();
    expect(label).toContain("converged");
    expect(label).not.toContain("duplicate");
  });
});
