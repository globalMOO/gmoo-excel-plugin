import { buildTemplateLayout, sheetRef, parseAddress, enumerateRangeAddresses } from "../excelService";

describe("sheetRef", () => {
  it("leaves simple names unquoted", () => {
    expect(sheetRef("Model")).toBe("Model");
    expect(sheetRef("Model_Def_1")).toBe("Model_Def_1");
  });
  it("quotes names with spaces or punctuation", () => {
    expect(sheetRef("My Model Def")).toBe("'My Model Def'");
    expect(sheetRef("A-B")).toBe("'A-B'");
  });
  it("escapes embedded single quotes", () => {
    expect(sheetRef("Bob's Sheet")).toBe("'Bob''s Sheet'");
  });
});

describe("buildTemplateLayout", () => {
  it("places inputs column-wise in column D starting at row 5", () => {
    const layout = buildTemplateLayout(3, 4, "Model");
    expect(layout.inputValueCol).toBe("D");
    expect(layout.inputDataStartRow).toBe(5);
    expect(layout.inputCells).toEqual([
      "Model!D5",
      "Model!D6",
      "Model!D7",
    ]);
  });

  it("places outcome formula cells in column B below the inputs", () => {
    const layout = buildTemplateLayout(3, 4, "Model");
    // 3 inputs → label row 9, header row 10, data starts row 11
    expect(layout.outcomeFormulaCol).toBe("B");
    expect(layout.outcomeDataStartRow).toBe(11);
    expect(layout.outputCells).toEqual([
      "Model!B11",
      "Model!B12",
      "Model!B13",
      "Model!B14",
    ]);
  });

  it("shifts the outcome block down as the input count grows", () => {
    const layout = buildTemplateLayout(8, 2, "Model");
    // 8 inputs occupy rows 5..12; outcome label at 14, header 15, data 16+
    expect(layout.outcomeDataStartRow).toBe(16);
    expect(layout.outputCells[0]).toBe("Model!B16");
  });

  it("quotes sheet names with spaces in generated addresses", () => {
    const layout = buildTemplateLayout(2, 1, "My Model Def");
    // 2 inputs → rows 5,6; outcome label 8, header 9, data starts 10.
    expect(layout.inputCells[0]).toBe("'My Model Def'!D5");
    expect(layout.outputCells[0]).toBe("'My Model Def'!B10");
  });

  it("round-trips its addresses through parseAddress", () => {
    const layout = buildTemplateLayout(3, 2, "My Model Def");
    const { sheet, cell } = parseAddress(layout.inputCells[0]);
    expect(sheet).toBe("My Model Def");
    expect(cell).toBe("D5");
    const out = parseAddress(layout.outputCells[1]);
    expect(out.sheet).toBe("My Model Def");
    expect(out.cell).toBe("B12");
  });
});

describe("enumerateRangeAddresses", () => {
  it("enumerates a single column top-to-bottom", () => {
    // B5:B7  → rowIndex 4, colIndex 1, 3 rows, 1 col
    expect(enumerateRangeAddresses("Sheet1", 4, 1, 3, 1)).toEqual([
      "Sheet1!B5",
      "Sheet1!B6",
      "Sheet1!B7",
    ]);
  });

  it("enumerates a single row left-to-right", () => {
    // B5:D5 → rowIndex 4, colIndex 1, 1 row, 3 cols
    expect(enumerateRangeAddresses("Sheet1", 4, 1, 1, 3)).toEqual([
      "Sheet1!B5",
      "Sheet1!C5",
      "Sheet1!D5",
    ]);
  });

  it("enumerates a rectangle row-major", () => {
    // A1:B2
    expect(enumerateRangeAddresses("S", 0, 0, 2, 2)).toEqual([
      "S!A1",
      "S!B1",
      "S!A2",
      "S!B2",
    ]);
  });

  it("quotes sheet names with spaces", () => {
    expect(enumerateRangeAddresses("My Sheet", 0, 0, 1, 1)).toEqual(["'My Sheet'!A1"]);
  });
});
