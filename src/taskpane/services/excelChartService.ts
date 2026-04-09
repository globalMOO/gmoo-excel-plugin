// Create native Excel charts via Office.js for final results
import type { Inverse } from "../types/gmoo";

export interface ChartData {
  iterations: Inverse[];
  inputVariableNames: string[];
  outcomeNames: string[];
}

export async function createResultsCharts(data: ChartData): Promise<void> {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;

    // Delete existing Results sheet if it exists
    const existingSheet = sheets.getItemOrNullObject("VSME Results");
    existingSheet.load("isNullObject");
    await context.sync();
    if (!existingSheet.isNullObject) {
      existingSheet.delete();
      await context.sync();
    }

    const sheet = sheets.add("VSME Results");
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

    // --- Chart layout: stacked vertically with generous spacing ---
    const chartWidth = 600;
    const chartHeight = 300;
    const rowsPerChart = 22; // ~300px ≈ 22 rows at default row height
    let chartRow = n + 4;

    // --- Error Convergence Chart ---
    const errorChart = sheet.charts.add(
      Excel.ChartType.line,
      sheet.getRange(`B1:B${n + 1}`),
      Excel.ChartSeriesBy.columns
    );
    errorChart.title.text = "Error Convergence (L1 Norm)";
    errorChart.setPosition("A" + chartRow);
    errorChart.height = chartHeight;
    errorChart.width = chartWidth;
    try { errorChart.series.getItemAt(0).setXAxisValues(sheet.getRange(`A2:A${n + 1}`)); } catch (_) { /* ignore */ }
    const allPositive = iterations.every((inv) => (inv.l1Norm ?? 0) > 0);
    if (allPositive) {
      try { errorChart.axes.getItem(Excel.ChartAxisType.value).logBase = 10; } catch (_) { /* not supported */ }
    }
    try { errorChart.legend.visible = false; } catch (_) { /* ignore */ }

    chartRow += rowsPerChart;

    // --- Input Variable Convergence Chart ---
    const inputDataStart = getColLetter(inputStartCol + 1); // E
    const inputDataEnd = getColLetter(inputStartCol + data.inputVariableNames.length);
    const inputChart = sheet.charts.add(
      Excel.ChartType.line,
      sheet.getRange(`${inputDataStart}1:${inputDataEnd}${n + 1}`),
      Excel.ChartSeriesBy.columns
    );
    inputChart.title.text = "Input Variable Convergence";
    inputChart.setPosition("A" + chartRow);
    inputChart.height = chartHeight;
    inputChart.width = chartWidth;
    try { inputChart.series.getItemAt(0).setXAxisValues(sheet.getRange(`D2:D${n + 1}`)); } catch (_) { /* ignore */ }

    chartRow += rowsPerChart;

    // --- Outcome Convergence Chart ---
    const outcomeDataStart = getColLetter(outcomeIterCol + 1);
    const outcomeDataEnd = getColLetter(outcomeIterCol + data.outcomeNames.length);
    const outcomeChart = sheet.charts.add(
      Excel.ChartType.line,
      sheet.getRange(`${outcomeDataStart}1:${outcomeDataEnd}${n + 1}`),
      Excel.ChartSeriesBy.columns
    );
    outcomeChart.title.text = "Outcome Convergence";
    outcomeChart.setPosition("A" + chartRow);
    outcomeChart.height = chartHeight;
    outcomeChart.width = chartWidth;
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
