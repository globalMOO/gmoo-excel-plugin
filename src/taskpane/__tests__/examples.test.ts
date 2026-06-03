import { EXAMPLES, getExampleById, isMultiSheetExample } from "../examples";

describe("polynomial-3x4 example", () => {
  const ex = getExampleById("polynomial-3x4");

  it("exists and is a simple-template example", () => {
    expect(ex).toBeDefined();
    expect(isMultiSheetExample(ex!)).toBe(false);
  });

  // Regression guard: the simple-template builder lays inputs out column-wise
  // with Current Value cells at D5/D6/D7 (see buildTemplateLayout). The example
  // formulas MUST reference those cells, not the old row-7 B/C/D layout.
  it("references the column-wise Current Value cells (D5/D6/D7)", () => {
    const formulas = ex!.setup.formulas!;
    const allText = Object.values(formulas).join(" ");
    // Must use the new D-column input cells…
    expect(allText).toMatch(/D5/);
    expect(allText).toMatch(/D6/);
    expect(allText).toMatch(/D7/);
    // …and must NOT reference the old row-7 layout (B7/C7).
    expect(allText).not.toMatch(/B7/);
    expect(allText).not.toMatch(/C7/);
  });

  it("has one formula per outcome", () => {
    const keys = Object.keys(ex!.setup.formulas!).sort();
    expect(keys).toEqual([...ex!.outcomeNames].sort());
  });
});

describe("EXAMPLES registry", () => {
  it("has unique ids", () => {
    const ids = EXAMPLES.map((e) => e.id);
    expect(new Set(ids).size).toBe(ids.length);
  });

  // Generic guard for EVERY simple-template (non-multi-sheet) example: the
  // template builder writes inputs into column D (Current Value). So every
  // formula must reference only column-D cells for its inputs — never columns
  // A/B/C (Name/Min/Max). This catches any new example hardcoded to the old
  // row-wise layout (inputs across row 7 in columns B/C/D).
  const simpleTemplateExamples = EXAMPLES.filter((e) => !isMultiSheetExample(e));

  it.each(simpleTemplateExamples.map((e) => [e.id, e] as const))(
    "simple-template example '%s' wires inputs from column D, not A/B/C",
    (_id, ex) => {
      const formulas = ex.setup.formulas!;
      expect(formulas).toBeDefined();
      const allText = Object.values(formulas).join(" ");
      // Inputs live in column D (Current Value). Formulas must read column D…
      expect(allText).toMatch(/\bD\d+\b/);
      // …and never columns A/B/C (Name/Min/Max), which is where the OLD
      // row-wise layout put inputs (B7/C7/D7 across row 7).
      expect(allText).not.toMatch(/\b[ABC]\d+\b/);
    }
  );
});
