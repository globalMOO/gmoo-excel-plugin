// Create native Excel charts via Office.js for final results
import type { Inverse } from "../types/gmoo";
import type { MultiSolveSolution } from "../hooks/useMultiSolve";
import type { InputVariable } from "../types/workbookState";

export interface ChartData {
  iterations: Inverse[];
  inputVariableNames: string[];
  outcomeNames: string[];
}

export async function createResultsCharts(data: ChartData): Promise<void> {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;

    // Delete existing Results sheet if it exists
    const existingSheet = sheets.getItemOrNullObject("GMOO Results");
    existingSheet.load("isNullObject");
    await context.sync();
    if (!existingSheet.isNullObject) {
      existingSheet.delete();
      await context.sync();
    }

    const sheet = sheets.add("GMOO Results");
    const iterations = data.iterations;
    const n = iterations.length;

    // --- Write backing data tables ---

    // Section 1: Error convergence (columns A-B) — A=Iteration, B=L1 Norm
    sheet.getRange("A1").values = [["Iteration"]];
    sheet.getRange("B1").values = [["L1 Norm"]];
    for (let i = 0; i < n; i++) {
      sheet.getRange(`A${i + 2}`).values = [[iterations[i].iteration]];
      sheet.getRange(`B${i + 2}`).values = [[iterations[i].l1Norm]];
    }

    // Section 2: Input variable convergence (starting column D)
    // D=Iteration (category labels), E+=variable values
    const inputStartCol = 4; // D
    sheet.getRange("D1").values = [["Iteration"]];
    for (let i = 0; i < n; i++) {
      sheet.getRange(`D${i + 2}`).values = [[iterations[i].iteration]];
    }
    for (let v = 0; v < data.inputVariableNames.length; v++) {
      const col = getColLetter(inputStartCol + 1 + v); // E, F, G, ...
      sheet.getRange(`${col}1`).values = [[data.inputVariableNames[v]]];
      for (let i = 0; i < n; i++) {
        sheet.getRange(`${col}${i + 2}`).values = [[iterations[i].input?.[v] ?? 0]];
      }
    }

    // Section 3: Outcome convergence (starting after input vars + gap)
    const outcomeIterCol = inputStartCol + 1 + data.inputVariableNames.length + 1;
    const outcomeIterLetter = getColLetter(outcomeIterCol);
    sheet.getRange(`${outcomeIterLetter}1`).values = [["Iteration"]];
    for (let i = 0; i < n; i++) {
      sheet.getRange(`${outcomeIterLetter}${i + 2}`).values = [[iterations[i].iteration]];
    }
    for (let o = 0; o < data.outcomeNames.length; o++) {
      const col = getColLetter(outcomeIterCol + 1 + o);
      sheet.getRange(`${col}1`).values = [[data.outcomeNames[o]]];
      for (let i = 0; i < n; i++) {
        sheet.getRange(`${col}${i + 2}`).values = [[iterations[i].output?.[o] ?? 0]];
      }
    }

    await context.sync();

    // --- Chart layout: absolute geometry (points) so charts never overlap ---
    // Row-offset positioning is fragile (chart height in px vs. variable row
    // heights), which let stacked charts collide. Anchor each chart at an
    // explicit top/left instead.
    const chartWidth = 600;
    const chartHeight = 300;
    const chartGap = 30;
    const baseTop = (n + 4) * 15; // ~15pt default row height, below the data
    const chartLeft = 12;
    const topFor = (index: number) => baseTop + index * (chartHeight + chartGap);

    const placeChart = (chart: Excel.Chart, index: number) => {
      chart.top = topFor(index);
      chart.left = chartLeft;
      chart.height = chartHeight;
      chart.width = chartWidth;
    };

    // --- Error Convergence Chart ---
    const errorChart = sheet.charts.add(
      Excel.ChartType.line,
      sheet.getRange(`B1:B${n + 1}`),
      Excel.ChartSeriesBy.columns
    );
    errorChart.title.text = "Error Convergence (L1 Norm)";
    placeChart(errorChart, 0);
    try { errorChart.series.getItemAt(0).setXAxisValues(sheet.getRange(`A2:A${n + 1}`)); } catch (_) { /* ignore */ }
    const allPositive = iterations.every((inv) => (inv.l1Norm ?? 0) > 0);
    if (allPositive) {
      try { errorChart.axes.getItem(Excel.ChartAxisType.value).logBase = 10; } catch (_) { /* not supported */ }
    }
    try { errorChart.legend.visible = false; } catch (_) { /* ignore */ }

    // --- Input Variable Convergence Chart ---
    const inputDataStart = getColLetter(inputStartCol + 1); // E
    const inputDataEnd = getColLetter(inputStartCol + data.inputVariableNames.length);
    const inputChart = sheet.charts.add(
      Excel.ChartType.line,
      sheet.getRange(`${inputDataStart}1:${inputDataEnd}${n + 1}`),
      Excel.ChartSeriesBy.columns
    );
    inputChart.title.text = "Input Variable Convergence";
    placeChart(inputChart, 1);
    try { inputChart.series.getItemAt(0).setXAxisValues(sheet.getRange(`D2:D${n + 1}`)); } catch (_) { /* ignore */ }

    // --- Outcome Convergence Chart ---
    const outcomeDataStart = getColLetter(outcomeIterCol + 1);
    const outcomeDataEnd = getColLetter(outcomeIterCol + data.outcomeNames.length);
    const outcomeChart = sheet.charts.add(
      Excel.ChartType.line,
      sheet.getRange(`${outcomeDataStart}1:${outcomeDataEnd}${n + 1}`),
      Excel.ChartSeriesBy.columns
    );
    outcomeChart.title.text = "Outcome Convergence";
    placeChart(outcomeChart, 2);
    try { outcomeChart.series.getItemAt(0).setXAxisValues(sheet.getRange(`${outcomeIterLetter}2:${outcomeIterLetter}${n + 1}`)); } catch (_) { /* ignore */ }

    sheet.activate();
    await context.sync();
  });
}

function getColLetter(colIndex: number): string {
  let result = "";
  let n = colIndex;
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

// --- Multi-Solve: native Excel radar charts ---

export interface MultiSolveRadarData {
  solutions: MultiSolveSolution[];
  inputVariables: InputVariable[];
  outcomeNames: string[];
}

/**
 * Export the Multi-Solve dual radar visualization to a native Excel sheet.
 *
 * Charts are built from RAW (native-unit) values so axis labels match what
 * the user expects — e.g. "$1,234" rather than "67%". This is a deliberate
 * divergence from the in-plugin Chart.js version: the sidebar has one narrow
 * polygon per chart and needs cross-axis comparability, which forces the
 * 0–100% normalization. The native Excel chart is larger, is viewed alone,
 * and users want true magnitudes.
 *
 * Both raw and normalized tables are written so the user can re-point the
 * chart source range to the normalized table if they want the percent view.
 */
export async function createMultiSolveRadarCharts(
  data: MultiSolveRadarData
): Promise<void> {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    const SHEET_NAME = "GMOO Multi-Solve Charts";

    const existing = sheets.getItemOrNullObject(SHEET_NAME);
    existing.load("isNullObject");
    await context.sync();
    if (!existing.isNullObject) {
      existing.delete();
      await context.sync();
    }

    const sheet = sheets.add(SHEET_NAME);
    const { solutions, inputVariables, outcomeNames } = data;
    const nSol = solutions.length;
    const nIn = inputVariables.length;
    const nOut = outcomeNames.length;

    if (nSol === 0) {
      sheet.getRange("A1").values = [["No Multi-Solve solutions to export."]];
      sheet.activate();
      await context.sync();
      return;
    }

    // Normalization helpers for the reference tables that sit below the raw ones.
    const outputMins = new Array<number>(nOut).fill(Number.POSITIVE_INFINITY);
    const outputMaxs = new Array<number>(nOut).fill(Number.NEGATIVE_INFINITY);
    for (const sol of solutions) {
      for (let i = 0; i < nOut; i++) {
        const v = sol.output?.[i];
        if (typeof v === "number" && isFinite(v)) {
          if (v < outputMins[i]) outputMins[i] = v;
          if (v > outputMaxs[i]) outputMaxs[i] = v;
        }
      }
    }

    const normalizeInput = (raw: number, v: InputVariable): number => {
      const range = v.max - v.min;
      if (range === 0) return 50;
      return Math.max(0, Math.min(1, (raw - v.min) / range)) * 100;
    };
    const normalizeOutput = (raw: number, idx: number): number => {
      const min = outputMins[idx];
      const max = outputMaxs[idx];
      if (!isFinite(min) || !isFinite(max) || min === max) return 50;
      return Math.max(0, Math.min(1, (raw - min) / (max - min))) * 100;
    };

    const solutionHeader = (prefix: string): (string | number)[] => {
      const header: (string | number)[] = [prefix];
      for (let s = 0; s < nSol; s++) header.push(`Solution ${s + 1}`);
      return header;
    };

    // --- Inputs — raw values (radar chart source) ---
    // Row 1: header. Rows 2..: one row per input variable with native values.
    sheet.getRange(`A1:${getColLetter(1 + nSol)}1`).values = [solutionHeader("Input")];
    for (let i = 0; i < nIn; i++) {
      const row: (string | number)[] = [inputVariables[i].name];
      for (let s = 0; s < nSol; s++) {
        row.push(solutions[s].input?.[i] ?? inputVariables[i].min);
      }
      sheet.getRange(`A${2 + i}:${getColLetter(1 + nSol)}${2 + i}`).values = [row];
    }
    sheet.getRange(`A1:${getColLetter(1 + nSol)}1`).format.font.bold = true;

    // --- Outputs — raw values (radar chart source) ---
    const outputHeaderRow = 2 + nIn + 2;
    sheet.getRange(`A${outputHeaderRow}:${getColLetter(1 + nSol)}${outputHeaderRow}`).values = [
      solutionHeader("Outcome"),
    ];
    for (let i = 0; i < nOut; i++) {
      const row: (string | number)[] = [outcomeNames[i]];
      for (let s = 0; s < nSol; s++) {
        row.push(solutions[s].output?.[i] ?? 0);
      }
      sheet.getRange(`A${outputHeaderRow + 1 + i}:${getColLetter(1 + nSol)}${outputHeaderRow + 1 + i}`).values = [row];
    }
    sheet.getRange(`A${outputHeaderRow}:${getColLetter(1 + nSol)}${outputHeaderRow}`).format.font.bold = true;

    // --- Reference: normalized (0–100%) tables below the raw ones ---
    // These aren't the chart source but are handy for comparing or for
    // rebuilding a percent-scale chart by changing the chart's source range.
    const normInputHeaderRow = outputHeaderRow + nOut + 3;
    sheet.getRange(`A${normInputHeaderRow}`).values = [["Inputs — normalized (% of min/max range)"]];
    sheet.getRange(`A${normInputHeaderRow}`).format.font.bold = true;
    sheet.getRange(`A${normInputHeaderRow + 1}:${getColLetter(1 + nSol)}${normInputHeaderRow + 1}`).values = [
      solutionHeader("Input"),
    ];
    for (let i = 0; i < nIn; i++) {
      const row: (string | number)[] = [inputVariables[i].name];
      for (let s = 0; s < nSol; s++) {
        row.push(normalizeInput(solutions[s].input?.[i] ?? inputVariables[i].min, inputVariables[i]));
      }
      sheet.getRange(
        `A${normInputHeaderRow + 2 + i}:${getColLetter(1 + nSol)}${normInputHeaderRow + 2 + i}`
      ).values = [row];
    }

    const normOutputHeaderRow = normInputHeaderRow + 2 + nIn + 2;
    sheet.getRange(`A${normOutputHeaderRow}`).values = [["Outputs — normalized (% of observed range)"]];
    sheet.getRange(`A${normOutputHeaderRow}`).format.font.bold = true;
    sheet.getRange(`A${normOutputHeaderRow + 1}:${getColLetter(1 + nSol)}${normOutputHeaderRow + 1}`).values = [
      solutionHeader("Outcome"),
    ];
    for (let i = 0; i < nOut; i++) {
      const row: (string | number)[] = [outcomeNames[i]];
      for (let s = 0; s < nSol; s++) {
        row.push(normalizeOutput(solutions[s].output?.[i] ?? 0, i));
      }
      sheet.getRange(
        `A${normOutputHeaderRow + 2 + i}:${getColLetter(1 + nSol)}${normOutputHeaderRow + 2 + i}`
      ).values = [row];
    }

    await context.sync();

    // --- Chart geometry (points) ---
    // Anchor both radars below the last reference table using explicit top/left
    // so they sit side by side and never overlap. Row-offset positioning let the
    // 360px-tall charts collide because chart height (px) ≠ variable row heights.
    const lastTableRow = normOutputHeaderRow + 1 + nOut;
    const chartTop = (lastTableRow + 2) * 15; // ~15pt default row height
    const chartHeight = 360;
    const chartWidth = 500;
    const chartGap = 24;

    // --- Inputs radar chart (raw values) ---
    const inputDataRange = sheet.getRange(
      `A1:${getColLetter(1 + nSol)}${1 + nIn}`
    );
    const inputChart = sheet.charts.add(
      Excel.ChartType.radar,
      inputDataRange,
      Excel.ChartSeriesBy.columns
    );
    inputChart.title.text = "Multi-Solve Inputs (native units)";
    inputChart.top = chartTop;
    inputChart.left = 12;
    inputChart.height = chartHeight;
    inputChart.width = chartWidth;

    // --- Outputs radar chart (raw values) ---
    const outputDataRange = sheet.getRange(
      `A${outputHeaderRow}:${getColLetter(1 + nSol)}${outputHeaderRow + nOut}`
    );
    const outputChart = sheet.charts.add(
      Excel.ChartType.radar,
      outputDataRange,
      Excel.ChartSeriesBy.columns
    );
    outputChart.title.text = "Multi-Solve Outputs (native units)";
    outputChart.top = chartTop;
    outputChart.left = 12 + chartWidth + chartGap;
    outputChart.height = chartHeight;
    outputChart.width = chartWidth;

    sheet.activate();
    await context.sync();
  });
}
