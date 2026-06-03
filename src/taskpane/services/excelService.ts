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
  // Contiguous mode (legacy template sheet) — all cells on one sheet in a grid.
  // Retained for backward compatibility; the current template builder populates
  // inputCells/outputCells instead and routes through the non-contiguous path.
  sheetName?: string;
  inputStartRow?: number;
  inputStartCol?: number;
  outputStartRow?: number;
  outputStartCol?: number;
  // Non-contiguous mode — each variable/outcome mapped to a specific cell.
  inputCells?: string[];  // full addresses like "Sheet1!D5", one per input variable
  outputCells?: string[]; // full addresses like "Sheet1!B10", one per outcome
}

/**
 * Geometry of the generated template sheet. Computed as a pure function so the
 * addresses can be unit-tested without an Excel context. The layout is
 * column-wise — each input variable is a row (Name | Min | Max | Current Value),
 * matching the Define Model table and the _GMOO_State sheet. The add-in writes
 * input values *down* the Current Value column (D); outcome formulas live in
 * column B with a live FORMULATEXT mirror in column C.
 *
 * Input data rows start at row 5 so a 3-input model lands its Current Value
 * cells at D5/D6/D7 — the simple-template examples' formulas depend on this.
 */
export interface TemplateLayout {
  titleRow: number;
  instructionRow: number;
  inputLabelRow: number;
  inputHeaderRow: number;
  inputDataStartRow: number;
  inputValueCol: string; // "D"
  outcomeLabelRow: number;
  outcomeHeaderRow: number;
  outcomeDataStartRow: number;
  outcomeFormulaCol: string; // "B"
  outcomeRefCol: string; // "C"
  inputCells: string[];
  outputCells: string[];
}

/** Wrap a sheet name in single quotes when it contains characters that require
 *  it in an A1 reference (spaces, etc). parseAddress() tolerates both forms. */
export function sheetRef(sheetName: string): string {
  return /^[A-Za-z0-9_]+$/.test(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`;
}

export function buildTemplateLayout(
  variableCount: number,
  outcomeCount: number,
  sheetName: string
): TemplateLayout {
  const titleRow = 1;
  const instructionRow = 2;
  const inputLabelRow = 3;
  const inputHeaderRow = 4;
  const inputDataStartRow = 5; // D5/D6/D7… — example formulas depend on this
  const inputValueCol = "D";

  const outcomeLabelRow = inputDataStartRow + variableCount + 1;
  const outcomeHeaderRow = outcomeLabelRow + 1;
  const outcomeDataStartRow = outcomeHeaderRow + 1;
  const outcomeFormulaCol = "B";
  const outcomeRefCol = "C";

  const ref = sheetRef(sheetName);
  const inputCells: string[] = [];
  for (let i = 0; i < variableCount; i++) {
    inputCells.push(`${ref}!${inputValueCol}${inputDataStartRow + i}`);
  }
  const outputCells: string[] = [];
  for (let i = 0; i < outcomeCount; i++) {
    outputCells.push(`${ref}!${outcomeFormulaCol}${outcomeDataStartRow + i}`);
  }

  return {
    titleRow,
    instructionRow,
    inputLabelRow,
    inputHeaderRow,
    inputDataStartRow,
    inputValueCol,
    outcomeLabelRow,
    outcomeHeaderRow,
    outcomeDataStartRow,
    outcomeFormulaCol,
    outcomeRefCol,
    inputCells,
    outputCells,
  };
}

const EXCEL_ERROR_VALUES = ["#VALUE!", "#REF!", "#NAME?", "#DIV/0!", "#NULL!", "#N/A", "#GETTING_DATA", "#NUM!"];

function isExcelError(value: unknown): boolean {
  if (typeof value === "string") {
    return EXCEL_ERROR_VALUES.some((e) => value.startsWith(e));
  }
  return false;
}

/** Apply a thin continuous box (outer edges + inner gridlines) to a range. */
function applyTableBorders(range: Excel.Range): void {
  const edges = [
    "EdgeTop",
    "EdgeBottom",
    "EdgeLeft",
    "EdgeRight",
    "InsideHorizontal",
    "InsideVertical",
  ] as const;
  for (const edge of edges) {
    const border = range.format.borders.getItem(edge as Excel.BorderIndex);
    border.style = Excel.BorderLineStyle.continuous;
    border.color = "#BFBFBF";
    border.weight = Excel.BorderWeight.thin;
  }
}

const INPUT_FILL = "#D9E1F2"; // light blue — cells the add-in writes
const OUTPUT_FILL = "#E2EFDA"; // light green — cells the user fills with formulas
const HEADER_FILL = "#4472C4"; // brand blue — table header rows
const HEADER_FONT = "#FFFFFF";

export async function createTemplateSheet(config: TemplateConfig): Promise<EvalConfig> {
  const layout = buildTemplateLayout(
    config.variables.length,
    config.outcomeNames.length,
    config.sheetName
  );

  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    const sheet = sheets.add(config.sheetName);
    const L = layout;

    // --- Title (row 1) ---
    const titleRange = sheet.getRange(`A${L.titleRow}:D${L.titleRow}`);
    titleRange.merge();
    titleRange.values = [[`GMOO Model: ${config.modelName}`, "", "", ""]];
    titleRange.format.font.bold = true;
    titleRange.format.font.size = 14;
    titleRange.format.font.color = "#1F3864";

    // --- Instructions (row 2) ---
    const instr = sheet.getRange(`A${L.instructionRow}:D${L.instructionRow}`);
    instr.merge();
    instr.values = [[
      "Enter each outcome's formula in the Formula column, referencing the Current Value cells (blue). The add-in fills Current Value automatically during evaluation.",
      "", "", "",
    ]];
    instr.format.font.size = 9;
    instr.format.font.italic = true;
    instr.format.font.color = "#605E5C";

    // --- Input Variables section ---
    sheet.getRange(`A${L.inputLabelRow}`).values = [["Input Variables"]];
    sheet.getRange(`A${L.inputLabelRow}`).format.font.bold = true;
    sheet.getRange(`A${L.inputLabelRow}`).format.font.color = "#1F3864";

    const inHeader = sheet.getRange(`A${L.inputHeaderRow}:D${L.inputHeaderRow}`);
    inHeader.values = [["Name", "Min", "Max", "Current Value"]];
    inHeader.format.font.bold = true;
    inHeader.format.font.color = HEADER_FONT;
    inHeader.format.fill.color = HEADER_FILL;

    for (let i = 0; i < config.variables.length; i++) {
      const row = L.inputDataStartRow + i;
      const v = config.variables[i];
      sheet.getRange(`A${row}`).values = [[v.name]];
      sheet.getRange(`B${row}`).values = [[v.min]];
      sheet.getRange(`C${row}`).values = [[v.max]];
      const cur = sheet.getRange(`${L.inputValueCol}${row}`);
      cur.values = [[0]]; // placeholder — add-in overwrites during evaluation
      cur.format.fill.color = INPUT_FILL;
    }
    applyTableBorders(
      sheet.getRange(
        `A${L.inputHeaderRow}:D${L.inputDataStartRow + config.variables.length - 1}`
      )
    );

    // --- Outcomes section ---
    sheet.getRange(`A${L.outcomeLabelRow}`).values = [["Outcomes"]];
    sheet.getRange(`A${L.outcomeLabelRow}`).format.font.bold = true;
    sheet.getRange(`A${L.outcomeLabelRow}`).format.font.color = "#1F3864";

    const outHeader = sheet.getRange(`A${L.outcomeHeaderRow}:C${L.outcomeHeaderRow}`);
    outHeader.values = [["Name", "Formula", "Reference (formula text)"]];
    outHeader.format.font.bold = true;
    outHeader.format.font.color = HEADER_FONT;
    outHeader.format.fill.color = HEADER_FILL;

    for (let i = 0; i < config.outcomeNames.length; i++) {
      const row = L.outcomeDataStartRow + i;
      sheet.getRange(`A${row}`).values = [[config.outcomeNames[i]]];
      const fcell = sheet.getRange(`${L.outcomeFormulaCol}${row}`);
      fcell.format.fill.color = OUTPUT_FILL;
      const formula = config.formulas?.[config.outcomeNames[i]];
      if (formula) {
        fcell.formulas = [[formula]];
      }
      // Live mirror of the formula text so the user can visually verify the
      // "path" each outcome is wired to (#N/A until a formula is entered).
      const refCell = sheet.getRange(`${L.outcomeRefCol}${row}`);
      refCell.formulas = [[`=FORMULATEXT(${L.outcomeFormulaCol}${row})`]];
      refCell.format.font.color = "#808080";
      refCell.format.font.italic = true;
    }
    applyTableBorders(
      sheet.getRange(
        `A${L.outcomeHeaderRow}:C${L.outcomeDataStartRow + config.outcomeNames.length - 1}`
      )
    );

    // --- Column widths & finish ---
    sheet.getRange("A:A").format.columnWidth = 150;
    sheet.getRange("B:B").format.columnWidth = 110;
    sheet.getRange("C:C").format.columnWidth = 160;
    sheet.getRange("D:D").format.columnWidth = 110;

    sheet.activate();
    await context.sync();

    return {
      sheetName: config.sheetName,
      variableCount: config.variables.length,
      outcomeCount: config.outcomeNames.length,
      inputCells: layout.inputCells,
      outputCells: layout.outputCells,
    };
  });
}

export async function writeInputValues(
  sheetName: string,
  row: number,
  startCol: number,
  values: number[]
): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    for (let i = 0; i < values.length; i++) {
      const colLetter = String.fromCharCode(64 + startCol + i);
      const cellAddr = `${colLetter}${row}`;
      sheet.getRange(cellAddr).values = [[values[i]]];
    }
    await context.sync();
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
        if (Number.isFinite(val)) {
          outputs.push(val);
        } else {
          errors.push(`Outcome ${i + 1}: Non-finite value (${val})`);
          outputs.push(0);
        }
      } else if (val === null || val === undefined || val === "") {
        errors.push(`Outcome ${i + 1}: Empty cell`);
        outputs.push(0);
      } else {
        const parsed = parseFloat(String(val));
        if (!Number.isFinite(parsed)) {
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
  if (config.inputCells && config.outputCells) {
    // Non-contiguous mode: write to individual cells, read from individual cells
    await writeCellValues(config.inputCells, inputValues);
    await calculateAndWait();
    return readCellValues(config.outputCells);
  }

  // Contiguous mode (template sheet)
  await writeInputValues(config.sheetName!, config.inputStartRow!, config.inputStartCol!, inputValues);
  await calculateAndWait();
  return readOutputValues(
    config.sheetName!,
    config.outputStartRow!,
    config.outputStartCol!,
    config.outcomeCount
  );
}

export async function evaluateAllCases(
  config: EvalConfig,
  inputCases: number[][],
  onProgress?: (current: number, total: number) => void
): Promise<{ outputCases: number[][]; errors: string[] }> {
  const outputCases: number[][] = [];
  const allErrors: string[] = [];

  if (!Array.isArray(inputCases) || inputCases.length === 0) {
    return { outputCases, errors: ["No input cases received — DOE payload is empty."] };
  }

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

/**
 * Enumerate the individual cell addresses of a rectangular range, row-major
 * (top→bottom, then left→right within each row). Pure so it can be unit-tested.
 * A single column reads top-to-bottom; a single row reads left-to-right.
 */
export function enumerateRangeAddresses(
  sheetName: string,
  startRowIndex: number, // 0-based
  startColIndex: number, // 0-based
  rowCount: number,
  columnCount: number
): string[] {
  const ref = sheetRef(sheetName);
  const addresses: string[] = [];
  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < columnCount; c++) {
      const col = columnLetter(startColIndex + c);
      const row = startRowIndex + r + 1; // A1 rows are 1-based
      addresses.push(`${ref}!${col}${row}`);
    }
  }
  return addresses;
}

/**
 * Read the currently-selected range and return the flattened, row-major list of
 * its individual cell addresses (sheet-qualified). Used to bulk-map inputs or
 * outputs in one gesture rather than picking each cell individually.
 */
export async function readSelectedCellAddresses(): Promise<string[]> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const ws = range.worksheet;
    range.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
    ws.load("name");
    await context.sync();
    return enumerateRangeAddresses(
      ws.name,
      range.rowIndex,
      range.columnIndex,
      range.rowCount,
      range.columnCount
    );
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

/**
 * Read the currently-selected range in Excel, returning both the absolute
 * address and the 2-D values array. Used by "Load from Excel selection"
 * buttons to pull configuration values out of any sheet — including a
 * different workbook via linked references.
 */
export async function readSelectedRangeWithValues(): Promise<{
  address: string;
  values: unknown[][];
}> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["address", "values"]);
    await context.sync();
    return { address: range.address, values: range.values };
  });
}

// --- Non-contiguous cell operations ---

export function parseAddress(fullAddress: string): { sheet: string; cell: string } {
  const idx = fullAddress.lastIndexOf("!");
  if (idx === -1) return { sheet: "", cell: fullAddress };
  return { sheet: fullAddress.slice(0, idx).replace(/'/g, ""), cell: fullAddress.slice(idx + 1) };
}

export async function writeCellValues(
  addresses: string[],
  values: number[]
): Promise<void> {
  await Excel.run(async (context) => {
    for (let i = 0; i < addresses.length; i++) {
      const { sheet, cell } = parseAddress(addresses[i]);
      const ws = sheet
        ? context.workbook.worksheets.getItem(sheet)
        : context.workbook.worksheets.getActiveWorksheet();
      ws.getRange(cell).values = [[values[i]]];
    }
    await context.sync();
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
        if (Number.isFinite(val)) {
          outputs.push(val);
        } else {
          errors.push(`Outcome ${i + 1} (${addresses[i]}): Non-finite value (${val})`);
          outputs.push(0);
        }
      } else if (val === null || val === undefined || val === "") {
        errors.push(`Outcome ${i + 1} (${addresses[i]}): Empty cell`);
        outputs.push(0);
      } else {
        const parsed = parseFloat(String(val));
        if (!Number.isFinite(parsed)) {
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

const STATE_SHEET_NAME = "_GMOO_State";

export interface GmooStateData {
  /** Project the cell mappings belong to. Optional for backward compat with
   *  pre-tagging sheets — when absent, callers may load tentatively. */
  projectId?: number;
  variables: { name: string; type: string; min: number; max: number; inputCell: string }[];
  outcomes: { name: string; outputCell: string }[];
}

/**
 * Check whether the _GMOO_State sheet exists in the workbook.
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
 * Read saved variable/outcome/cell-mapping state from the _GMOO_State sheet.
 * Returns null if the sheet doesn't exist or can't be parsed.
 */
export async function loadStateSheet(): Promise<GmooStateData | null> {
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

    // Find the "Variables", "Outcomes", and optional "Project ID" rows.
    let varHeaderRow = -1;
    let outHeaderRow = -1;
    let projectId: number | undefined;
    for (let r = 0; r < rows.length; r++) {
      const label = String(rows[r][0]).trim();
      if (label === "Variables") varHeaderRow = r;
      else if (label === "Outcomes") outHeaderRow = r;
      else if (label === "Project ID") {
        const raw = rows[r][1];
        const n = typeof raw === "number" ? raw : parseInt(String(raw), 10);
        if (!isNaN(n) && n > 0) projectId = n;
      }
    }

    if (varHeaderRow === -1 || outHeaderRow === -1) return null;

    // Parse variables: rows between varHeaderRow+1 and outHeaderRow (blank row before Outcomes)
    const variables: GmooStateData["variables"] = [];
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
    const outcomes: GmooStateData["outcomes"] = [];
    for (let r = outHeaderRow + 1; r < rows.length; r++) {
      const name = String(rows[r][1] ?? "").trim();
      if (!name) continue;
      outcomes.push({
        name,
        outputCell: String(rows[r][2] ?? "").trim(),
      });
    }

    if (variables.length === 0 || outcomes.length === 0) return null;
    return { projectId, variables, outcomes };
  });
}

/**
 * Save variable/outcome/cell-mapping state to the _GMOO_State sheet.
 * Creates or replaces the sheet.
 */
export async function saveStateSheet(data: GmooStateData): Promise<void> {
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
    sheet.getRange("A1").values = [["GMOO Configuration"]];
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A1").format.font.size = 12;

    // Optional project tag (row 2) — used by App.tsx to verify the saved
    // mappings still describe the currently-active project before auto-loading.
    if (typeof data.projectId === "number" && data.projectId > 0) {
      sheet.getRange("A2").values = [["Project ID"]];
      sheet.getRange("A2").format.font.color = "#605E5C";
      sheet.getRange("B2").values = [[data.projectId]];
    }

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

// --- Multi-Solve solution persistence ---

const MULTI_SOLVE_SHEET = "GMOO Multi-Solve";

export interface MultiSolveRunRow {
  runIndex: number;        // 0-based
  status: "unique" | "failed";
  solutionIdx: number | null;   // 1-based solution index when converged; null on failure
  distanceToNearest: number | null; // normalized Euclidean distance to the nearest already-collected solution (diagnostic only; null for the first one)
  satisfied: boolean;
  l1Norm: number | null;        // null on failure
  iterations: number | null;    // null on failure
  initialInputs: number[];
  finalInputs: number[] | null;
  outputs: number[] | null;
  note?: string;                // optional diagnostic message (e.g. error reason for failed runs)
}

/**
 * Appends a per-run diagnostic row to the "GMOO Multi-Solve" sheet. Every
 * random-start run writes one row — unique, duplicate, and failed alike —
 * so the user can diagnose why the optimizer does or doesn't find distinct
 * solutions. `resetSheet=true` wipes any existing sheet (called on the first
 * row of a new Multi-Solve run).
 */
export async function writeMultiSolveRun(
  row: MultiSolveRunRow,
  inputVarNames: string[],
  outcomeNames: string[],
  resetSheet: boolean
): Promise<void> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;

    const existing = sheets.getItemOrNullObject(MULTI_SOLVE_SHEET);
    existing.load("name");
    await context.sync();

    let sheet: Excel.Worksheet;
    const headerRow = [
      "Run",
      "Status",
      "Solution #",
      "Dist. to nearest",
      "Satisfied",
      "L1 Norm",
      "Iter",
      ...inputVarNames.map((n) => `init: ${n}`),
      ...inputVarNames.map((n) => `final: ${n}`),
      ...outcomeNames.map((n) => `out: ${n}`),
      "Note",
    ];

    const wasNull = existing.isNullObject;
    if (wasNull) {
      sheet = sheets.add(MULTI_SOLVE_SHEET);
      writeHeaderRow(sheet, headerRow);
    } else if (resetSheet) {
      existing.delete();
      await context.sync();
      sheet = sheets.add(MULTI_SOLVE_SHEET);
      writeHeaderRow(sheet, headerRow);
    } else {
      sheet = existing;
    }

    let nextRow = 2;
    if (!resetSheet && !wasNull) {
      const used = sheet.getUsedRange(true);
      used.load("rowCount");
      await context.sync();
      nextRow = (used.rowCount || 1) + 1;
    }

    const nInputs = inputVarNames.length;
    const nOutputs = outcomeNames.length;
    // Pad arrays so every row has the same number of columns — Excel won't
    // accept a jagged range assignment.
    const padded = (arr: number[] | null, n: number): (number | string)[] =>
      arr ? arr.slice(0, n) : new Array(n).fill("");

    const rowValues: (string | number | boolean)[] = [
      row.runIndex + 1,
      row.status,
      row.solutionIdx ?? "",
      row.distanceToNearest !== null ? row.distanceToNearest : "",
      row.status === "failed" ? "" : (row.satisfied ? "Yes" : "No"),
      row.l1Norm !== null ? row.l1Norm : "",
      row.iterations !== null ? row.iterations : "",
      ...padded(row.initialInputs, nInputs),
      ...padded(row.finalInputs, nInputs),
      ...padded(row.outputs, nOutputs),
      row.note ?? "",
    ];

    const endColLetter = columnLetter(rowValues.length - 1);
    const targetRange = sheet.getRange(`A${nextRow}:${endColLetter}${nextRow}`);
    targetRange.values = [rowValues];

    await context.sync();
  });
}

function writeHeaderRow(sheet: Excel.Worksheet, headers: string[]): void {
  const endCol = columnLetter(headers.length - 1);
  const headerRange = sheet.getRange(`A1:${endCol}1`);
  headerRange.values = [headers];
  headerRange.format.font.bold = true;
  headerRange.format.fill.color = "#D9E1F2";
}

function columnLetter(zeroBasedIndex: number): string {
  let n = zeroBasedIndex;
  let s = "";
  while (n >= 0) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}
