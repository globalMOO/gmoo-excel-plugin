import { randomInputWithinBounds } from "../sampling";
import type { InputVariable } from "../../types/workbookState";

const floatVar = (min: number, max: number): InputVariable => ({ name: "f", type: "float", min, max });
const intVar = (min: number, max: number): InputVariable => ({ name: "i", type: "integer", min, max });
const boolVar = (): InputVariable => ({ name: "b", type: "boolean", min: 0, max: 1 });

describe("randomInputWithinBounds", () => {
  it("returns one value per variable", () => {
    const out = randomInputWithinBounds([floatVar(0, 1), floatVar(2, 3)]);
    expect(out).toHaveLength(2);
  });

  it("keeps float draws within [min, max] across many samples", () => {
    const v = floatVar(-5, 10);
    for (let i = 0; i < 500; i++) {
      const [x] = randomInputWithinBounds([v]);
      expect(x).toBeGreaterThanOrEqual(-5);
      expect(x).toBeLessThanOrEqual(10);
    }
  });

  it("produces integers within the inclusive range", () => {
    const v = intVar(1, 4);
    for (let i = 0; i < 500; i++) {
      const [x] = randomInputWithinBounds([v]);
      expect(Number.isInteger(x)).toBe(true);
      expect(x).toBeGreaterThanOrEqual(1);
      expect(x).toBeLessThanOrEqual(4);
    }
  });

  it("produces only 0 or 1 for boolean", () => {
    for (let i = 0; i < 200; i++) {
      const [x] = randomInputWithinBounds([boolVar()]);
      expect([0, 1]).toContain(x);
    }
  });

  it("returns the floor when an integer range is degenerate", () => {
    expect(randomInputWithinBounds([intVar(3, 3)])).toEqual([3]);
    expect(randomInputWithinBounds([intVar(5, 4)])).toEqual([5]);
  });
});
