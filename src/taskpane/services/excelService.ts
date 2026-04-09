// Office.js Excel operations for template sheet creation and case evaluation
import type { InputVariable } from "../types/workbookState";

const RECALC_POLL_INTERVAL = 50;
const RECALC_TIMEOUT = 30000;

export interface TemplateConfig {
  modelName: string;
  variables: InputVariable[];
  outcomeNames: string[];
  sheetName: string;
  /** Optional formulas keyed by outcome name (e.g. { W: "=B7^2+2*C7" }) */
  formulas?: Record<string, string>;
}

export interface EvalConfig {
  variableCount: number;
  outcomeCount: number;
  // Contiguous mode (template sheet) — all cells on one sheet in a grid
  sheetName?: string;
  inputStartRow?: number;
  inputStartCol?: number;
  outputStartRow?: number;
  outputStartCol?: number;
  // Non-contiguous mode (existing sheet) — each variable mapped to a specific cell
  inputCells?: string[];  // full addresses like "Sheet1!B7", one per input variable
  outputCells?: string[]; // full addresses like "Sheet1!B11", one per outcome
}

const EXCEL_ERROR_VALUES = ["#VALUE!", "#REF!", "#NAME?", "#DIV/0!", "#NULL!", "#N/A", "#GETTING_DATA", "#NUM!"];

function isExcelError(value: unknown): boolean {
  if (typeof value === "string") {
    return EXCEL_ERROR_VALUES.some((e) => value.startsWith(e));
  }
  return false;
}

export async function createTemplateSheet(config: TemplateConfig): Promise<EvalConfig> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    const sheet = sheets.add(config.sheetName);

    // Row 1: Model name header
    const headerRange = sheet.getRange("A1:D1");
    headerRange.merge();
    headerRange.values = [[`VSME Model: ${config.modelName}`, "", "", ""]];
    headerRange.format.font.bold = true;
    headerRange.format.font.size = 14;

    // Row 3: Input Variables header
    sheet.getRange("A3").values = [["Input Variables"]];
    sheet.getRange("A3").format.font.bold = true;

    // Row 4: Variable names across columns B+
    // Row 5: Min values
    // Row 6: Max values
    // Row 7: Current Value cells (add-in writes here during evaluation)
    sheet.getRange("A4").values = [["Name"]];
    sheet.getRange("A5").values = [["Min"]];
    sheet.getRange("A6").values = [["Max"]];
    sheet.getRange("A7").values = [["Current Value"]];

    for (let i = 0; i < config.variables.length; i++) {
      const col = String.fromCharCode(66 + i); // B, C, D, ...
      sheet.getRange(`${col}4`).values = [[config.variables[i].name]];
      sheet.getRange(`${col}5`).values = [[config.variables[i].min]];
      sheet.getRange(`${col}6`).values = [[config.variables[i].max]];
      sheet.getRange(`${col}7`).values = [[0]]; // placeholder
      sheet.getRange(`${col}7`).format.fill.color = "#D9E1F2";
    }

    // Row 9: Outcomes header
    sheet.getRange("A9").values = [["Outcomes"]];
    sheet.getRange("A9").format.font.bold = true;
    sheet.getRange("A10").values = [["Name"]];
    sheet.getRange("B10").values = [["Formula"]];

    // Row 11+: Outcome names in col A, formula cells in col B
    for (let i = 0; i < config.outcomeNames.length; i++) {
      const row = 11 + i;
      sheet.getRange(`A${row}`).values = [[config.outcomeNames[i]]];
      sheet.getRange(`B${row}`).format.fill.color = "#E2EFDA";
      const formula = config.formulas?.[config.outcomeNames[i]];
      if (formula) {
        sheet.getRange(`B${row}`).formulas = [[formula]];
      }
    }

    // Auto-fit columns
    sheet.getUsedRange().format.autofitColumns();

    sheet.activate();
    await context.sync();

    return {
      sheetName: config.sheetName,
      inputStartRow: 7,
      inputStartCol: 2, // column B = 2 (1-indexed)
      outputStartRow: 11,
      outputStartCol: 2, // column B
      variableCount: config.variables.length,
      outcomeCount: config.outcomeNames.length,
    };
  });
}

export async function writeInputValues(
  sheetName: string,
  row: number,
  startCol: number,
  values: number[]
): Promise<void> {
  console.log(`[Excel] writeInputValues: sheet="${sheetName}", row=${row}, startCol=${startCol}, values=`, values);
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    for (let i = 0; i < values.length; i++) {
      const colLetter = String.fromCharCode(64 + startCol + i);
      const cellAddr = `${colLetter}${row}`;
      sheet.getRange(cellAddr).values = [[values[i]]];
    }
    await context.sync();
    console.log(`[Excel] writeInputValues: sync complete`);
  });
}

export async function calculateAndWait(): Promise<void> {
  await Excel.run(async (context) => {
    context.workbook.application.calculate(Excel.CalculationType.full);
    await context.sync();
  });

  // Poll calculationState until done
  const start = Date.now();
  while (Date.now() - start < RECALC_TIMEOUT) {
    const state = await Excel.run(async (context) => {
      const app = context.workbook.application;
      app.load("calculationState");
      await context.sync();
      return app.calculationState;
    });

    if (state === Excel.CalculationState.done) return;
    await delay(RECALC_POLL_INTERVAL);
  }
  throw new Error("Calculation timed out after 30 seconds.");
}

export async function readOutputValues(
  sheetName: string,
  startRow: number,
  col: number,
  count: number
): Promise<{ outputs: number[]; errors: string[] }> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const colLetter = String.fromCharCode(64 + col);
    const range = sheet.getRange(`${colLetter}${startRow}:${colLetter}${startRow + count - 1}`);
    range.load("values");
    await context.sync();

    const outputs: number[] = [];
    const errors: string[] = [];

    for (let i = 0; i < count; i++) {
      const val = range.values[i][0];
      if (isExcelError(val)) {
        errors.push(`Outcome ${i + 1}: ${val}`);
        outputs.push(0);
      } else if (typeof val === "number") {
        outputs.push(val);
      } else {
        const parsed = parseFloat(String(val));
        if (isNaN(parsed)) {
          errors.push(`Outcome ${i + 1}: Non-numeric value "${val}"`);
          outputs.push(0);
        } else {
          outputs.push(parsed);
        }
      }
    }

    return { outputs, errors };
  });
}

export async function evaluateCase(
  config: EvalConfig,
  inputValues: number[]
): Promise<{ outputs: number[]; errors: string[] }> {
  console.log(`[Excel] evaluateCase: config=`, config, `inputs=`, inputValues);

  if (config.inputCells && config.outputCells) {
    // Non-contiguous mode: write to individual cells, read from individual cells
    await writeCellValues(config.inputCells, inputValues);
    await calculateAndWait();
    const result = await readCellValues(config.outputCells);
    console.log(`[Excel] evaluateCase: outputs=`, result.outputs, `errors=`, result.errors);
    return result;
  }

  // Contiguous mode (template sheet)
  await writeInputValues(config.sheetName!, config.inputStartRow!, config.inputStartCol!, inputValues);
  await calculateAndWait();
  const result = await readOutputValues(
    config.sheetName!,
    config.outputStartRow!,
    config.outputStartCol!,
    config.outcomeCount
  );

  console.log(`[Excel] evaluateCase: outputs=`, result.outputs, `errors=`, result.errors);
  return result;
}

export async function evaluateAllCases(
  config: EvalConfig,
  inputCases: number[][],
  onProgress?: (current: number, total: number) => void
): Promise<{ outputCases: number[][]; errors: string[] }> {
  const outputCases: number[][] = [];
  const allErrors: string[] = [];

  // Suspend screen updating for performance
  await Excel.run(async (context) => {
    context.workbook.application.suspendScreenUpdatingUntilNextSync();
    await context.sync();
  });

  for (let i = 0; i < inputCases.length; i++) {
    onProgress?.(i + 1, inputCases.length);
    const result = await evaluateCase(config, inputCases[i]);
    outputCases.push(result.outputs);
    if (result.errors.length > 0) {
      allErrors.push(`Case ${i + 1}: ${result.errors.join(", ")}`);
    }
  }

  return { outputCases, errors: allErrors };
}

export async function readSelectedRange(): Promise<string> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    return range.address;
  });
}

export async function readRangeValues(address: string): Promise<unknown[][]> {
  return Excel.run(async (context) => {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
    range.load("values");
    await context.sync();
    return range.values;
  });
}

// --- Non-contiguous cell operations ---

function parseAddress(fullAddress: string): { sheet: string; cell: string } {
  const idx = fullAddress.lastIndexOf("!");
  if (idx === -1) return { sheet: "", cell: fullAddress };
  return { sheet: fullAddress.slice(0, idx).replace(/'/g, ""), cell: fullAddress.slice(idx + 1) };
}

export async function writeCellValues(
  addresses: string[],
  values: number[]
): Promise<void> {
  console.log(`[Excel] writeCellValues: ${addresses.length} cells`);
  await Excel.run(async (context) => {
    for (let i = 0; i < addresses.length; i++) {
      const { sheet, cell } = parseAddress(addresses[i]);
      const ws = sheet
        ? context.workbook.worksheets.getItem(sheet)
        : context.workbook.worksheets.getActiveWorksheet();
      ws.getRange(cell).values = [[values[i]]];
    }
    await context.sync();
    console.log(`[Excel] writeCellValues: sync complete`);
  });
}

export async function readCellValues(
  addresses: string[]
): Promise<{ outputs: number[]; errors: string[] }> {
  return Excel.run(async (context) => {
    const ranges: Excel.Range[] = [];
    for (const addr of addresses) {
      const { sheet, cell } = parseAddress(addr);
      const ws = sheet
        ? context.workbook.worksheets.getItem(sheet)
        : context.workbook.worksheets.getActiveWorksheet();
      const range = ws.getRange(cell);
      range.load("values");
      ranges.push(range);
    }
    await context.sync();

    const outputs: number[] = [];
    const errors: string[] = [];

    for (let i = 0; i < ranges.length; i++) {
      const val = ranges[i].values[0][0];
      if (isExcelError(val)) {
        errors.push(`Outcome ${i + 1} (${addresses[i]}): ${val}`);
        outputs.push(0);
      } else if (typeof val === "number") {
        outputs.push(val);
      } else {
        const parsed = parseFloat(String(val));
        if (isNaN(parsed)) {
          errors.push(`Outcome ${i + 1} (${addresses[i]}): Non-numeric value "${val}"`);
          outputs.push(0);
        } else {
          outputs.push(parsed);
        }
      }
    }

    return { outputs, errors };
  });
}

// --- Multi-sheet example creation ---

import type { ExampleSheet } from "../examples";

/**
 * Creates multiple Excel sheets from an example definition.
 * Returns an EvalConfig using non-contiguous cell mapping.
 */
export async function createExampleSheets(
  sheets: ExampleSheet[],
  inputCells: string[],
  outputCells: string[],
  variableCount: number,
  outcomeCount: number
): Promise<EvalConfig> {
  await Excel.run(async (context) => {
    for (const sheetDef of sheets) {
      const sheet = context.workbook.worksheets.add(sheetDef.name);

      // Write data cell by cell (handles mixed values and formulas)
      for (let r = 0; r < sheetDef.data.length; r++) {
        const row = sheetDef.data[r];
        for (let c = 0; c < row.length; c++) {
          const val = row[c];
          if (val === null || val === undefined) continue;
          const colLetter = String.fromCharCode(65 + c);
          const cellAddr = `${colLetter}${r + 1}`;
          const cell = sheet.getRange(cellAddr);
          if (typeof val === "string" && val.startsWith("=")) {
            cell.formulas = [[val]];
          } else {
            cell.values = [[val]];
          }
        }
      }

      // Merge ranges
      if (sheetDef.merges) {
        for (const merge of sheetDef.merges) {
          sheet.getRange(merge).merge();
        }
      }

      // Bold rows
      if (sheetDef.boldRows) {
        for (const rowIdx of sheetDef.boldRows) {
          const lastCol = Math.max(...sheetDef.data.map((r) => r.length));
          const endColLetter = String.fromCharCode(64 + lastCol);
          sheet.getRange(`A${rowIdx + 1}:${endColLetter}${rowIdx + 1}`).format.font.bold = true;
        }
      }

      // Column widths
      if (sheetDef.columnWidths) {
        for (let c = 0; c < sheetDef.columnWidths.length; c++) {
          const w = sheetDef.columnWidths[c];
          if (w != null) {
            const colLetter = String.fromCharCode(65 + c);
            sheet.getRange(`${colLetter}:${colLetter}`).format.columnWidth = w * 7; // chars → pixels approx
          }
        }
      }

      // Highlight input cells
      if (sheetDef.inputHighlights) {
        for (const cellAddr of sheetDef.inputHighlights) {
          sheet.getRange(cellAddr).format.fill.color = "#D9E1F2"; // light blue
        }
      }
    }

    // Activate the Model sheet (second in the array) or fall back to last sheet
    if (sheets.length > 1) {
      context.workbook.worksheets.getItem(sheets[1].name).activate();
    } else if (sheets.length > 0) {
      context.workbook.worksheets.getItem(sheets[0].name).activate();
    }

    await context.sync();
  });

  return {
    variableCount,
    outcomeCount,
    inputCells,
    outputCells,
  };
}

// --- State sheet persistence ---

const STATE_SHEET_NAME = "_VSME_State";

export interface VsmeStateData {
  variables: { name: string; type: string; min: number; max: number; inputCell: string }[];
  outcomes: { name: string; outputCell: string }[];
}

/**
 * Check whether the _VSME_State sheet exists in the workbook.
 */
export async function hasStateSheet(): Promise<boolean> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    return sheets.items.some((s) => s.name === STATE_SHEET_NAME);
  });
}

/**
 * Read saved variable/outcome/cell-mapping state from the _VSME_State sheet.
 * Returns null if the sheet doesn't exist or can't be parsed.
 */
export async function loadStateSheet(): Promise<VsmeStateData | null> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    if (!sheets.items.some((s) => s.name === STATE_SHEET_NAME)) return null;

    const sheet = sheets.getItem(STATE_SHEET_NAME);
    const usedRange = sheet.getUsedRange();
    usedRange.load("values");
    await context.sync();

    const rows = usedRange.values as (string | number)[][];

    // Find the "Variables" and "Outcomes" section headers
    let varHeaderRow = -1;
    let outHeaderRow = -1;
    for (let r = 0; r < rows.length; r++) {
      const label = String(rows[r][0]).trim();
      if (label === "Variables") varHeaderRow = r;
      if (label === "Outcomes") outHeaderRow = r;
    }

    if (varHeaderRow === -1 || outHeaderRow === -1) return null;

    // Parse variables: rows between varHeaderRow+1 and outHeaderRow (blank row before Outcomes)
    const variables: VsmeStateData["variables"] = [];
    for (let r = varHeaderRow + 1; r < outHeaderRow; r++) {
      const name = String(rows[r][1] ?? "").trim();
      if (!name) continue; // skip blank rows
      variables.push({
        name,
        type: String(rows[r][2] ?? "float"),
        min: Number(rows[r][3]) || 0,
        max: Number(rows[r][4]) || 0,
        inputCell: String(rows[r][5] ?? "").trim(),
      });
    }

    // Parse outcomes: rows after outHeaderRow until end
    const outcomes: VsmeStateData["outcomes"] = [];
    for (let r = outHeaderRow + 1; r < rows.length; r++) {
      const name = String(rows[r][1] ?? "").trim();
      if (!name) continue;
      outcomes.push({
        name,
        outputCell: String(rows[r][2] ?? "").trim(),
      });
    }

    if (variables.length === 0 || outcomes.length === 0) return null;
    return { variables, outcomes };
  });
}

/**
 * Save variable/outcome/cell-mapping state to the _VSME_State sheet.
 * Creates or replaces the sheet.
 */
export async function saveStateSheet(data: VsmeStateData): Promise<void> {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    // Delete existing state sheet if present
    const existing = sheets.items.find((s) => s.name === STATE_SHEET_NAME);
    if (existing) {
      existing.delete();
      await context.sync();
    }

    const sheet = sheets.add(STATE_SHEET_NAME);

    // Header
    sheet.getRange("A1").values = [["VSME Configuration"]];
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A1").format.font.size = 12;

    // Variables section
    const varHeaderRow = 3;
    sheet.getRange(`A${varHeaderRow}`).values = [["Variables"]];
    sheet.getRange(`A${varHeaderRow}`).format.font.bold = true;
    sheet.getRange(`B${varHeaderRow}`).values = [["Name"]];
    sheet.getRange(`C${varHeaderRow}`).values = [["Type"]];
    sheet.getRange(`D${varHeaderRow}`).values = [["Min"]];
    sheet.getRange(`E${varHeaderRow}`).values = [["Max"]];
    sheet.getRange(`F${varHeaderRow}`).values = [["Input Cell"]];
    sheet.getRange(`B${varHeaderRow}:F${varHeaderRow}`).format.font.bold = true;

    for (let i = 0; i < data.variables.length; i++) {
      const row = varHeaderRow + 1 + i;
      const v = data.variables[i];
      sheet.getRange(`B${row}`).values = [[v.name]];
      sheet.getRange(`C${row}`).values = [[v.type]];
      sheet.getRange(`D${row}`).values = [[v.min]];
      sheet.getRange(`E${row}`).values = [[v.max]];
      sheet.getRange(`F${row}`).values = [[v.inputCell]];
      // Highlight the cell address
      sheet.getRange(`F${row}`).format.font.color = "#0078D4";
    }

    // Outcomes section (2 rows after last variable)
    const outHeaderRow = varHeaderRow + 1 + data.variables.length + 1;
    sheet.getRange(`A${outHeaderRow}`).values = [["Outcomes"]];
    sheet.getRange(`A${outHeaderRow}`).format.font.bold = true;
    sheet.getRange(`B${outHeaderRow}`).values = [["Name"]];
    sheet.getRange(`C${outHeaderRow}`).values = [["Output Cell"]];
    sheet.getRange(`B${outHeaderRow}:C${outHeaderRow}`).format.font.bold = true;

    for (let i = 0; i < data.outcomes.length; i++) {
      const row = outHeaderRow + 1 + i;
      const o = data.outcomes[i];
      sheet.getRange(`B${row}`).values = [[o.name]];
      sheet.getRange(`C${row}`).values = [[o.outputCell]];
      sheet.getRange(`C${row}`).format.font.color = "#0078D4";
    }

    // Column widths
    sheet.getRange("A:A").format.columnWidth = 90;
    sheet.getRange("B:B").format.columnWidth = 120;
    sheet.getRange("C:C").format.columnWidth = 100;
    sheet.getRange("F:F").format.columnWidth = 120;

    // Don't activate — keep user on their current sheet
    await context.sync();
  });
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
