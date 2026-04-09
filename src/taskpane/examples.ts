// Pre-built example problems for guiding new users
import type { InputVariable } from "./types/workbookState";
import { InputType, ObjectiveType } from "./types/gmoo";

export interface ExampleObjective {
  type: ObjectiveType;
  target: string;
  minBound: string;
  maxBound: string;
}

/**
 * A sheet to be created in Excel as part of an example.
 * Each cell can be a string (formula if starts with "="), number, or null (empty).
 */
export interface ExampleSheet {
  name: string;
  /** 2D array of cell data, row-major. Strings starting with "=" are formulas. */
  data: (string | number | null)[][];
  /** 0-based row indices to format as bold */
  boldRows?: number[];
  /** Cell addresses (sheet-local, e.g. "B5") to highlight as input cells */
  inputHighlights?: string[];
  /** Merge ranges, e.g. ["A1:E1"] */
  merges?: string[];
  /** Column widths in characters (index matches column: 0=A, 1=B, ...) */
  columnWidths?: (number | null)[];
}

/**
 * Setup configuration carried by an example.
 * Simple examples use `formulas` (single template sheet).
 * Complex examples use `sheets` + `inputCells` + `outputCells` (multi-sheet, non-contiguous).
 */
export interface ExampleSetup {
  /** Simple mode: formulas keyed by outcome name, inserted into the auto-generated template sheet */
  formulas?: Record<string, string>;
  /** Complex mode: full sheet definitions to create in Excel */
  sheets?: ExampleSheet[];
  /** Complex mode: full cell addresses for each input variable, e.g. "Model!B5" */
  inputCells?: string[];
  /** Complex mode: full cell addresses for each output, e.g. "Model!B8" */
  outputCells?: string[];
  /** Instructions shown to the user if the example needs explanation */
  setupInstructions?: string;
}

export interface Example {
  id: string;
  name: string;
  description: string;
  /** Difficulty/complexity hint */
  complexity?: "beginner" | "intermediate" | "advanced";
  variables: InputVariable[];
  outcomeNames: string[];
  /** Setup configuration — formulas, sheets, cell mappings */
  setup: ExampleSetup;
  /** Default objective for each outcome (same order as outcomeNames) */
  objectives: ExampleObjective[];
}

/** Whether this example uses multi-sheet mode (vs simple template) */
export function isMultiSheetExample(example: Example): boolean {
  return !!(example.setup.sheets && example.setup.inputCells && example.setup.outputCells);
}

// ---------------------------------------------------------------------------
// Example definitions
// ---------------------------------------------------------------------------

export const EXAMPLES: Example[] = [
  // --- Beginner: simple polynomial ---
  {
    id: "polynomial-3x4",
    name: "Polynomial (3 inputs, 4 outputs)",
    description: "Simple polynomial functions on a single sheet — good first example to verify the workflow.",
    complexity: "beginner",
    variables: [
      { name: "A", type: InputType.Float, min: 0, max: 100 },
      { name: "B", type: InputType.Float, min: 0, max: 50 },
      { name: "C", type: InputType.Float, min: -10, max: 10 },
    ],
    outcomeNames: ["W", "X", "Y", "Z"],
    setup: {
      formulas: {
        W: "=B7^2+2*C7-D7",
        X: "=B7*C7+D7^2",
        Y: "=SIN(B7)+C7*D7",
        Z: "=B7+C7^2-3*D7",
      },
    },
    // Truth case: A=12, B=34, C=-5.6
    objectives: [
      { type: ObjectiveType.Percent, target: "217.6", minBound: "-5", maxBound: "5" },
      { type: ObjectiveType.Percent, target: "439.36", minBound: "-5", maxBound: "5" },
      { type: ObjectiveType.Percent, target: "-190.9366", minBound: "-5", maxBound: "5" },
      { type: ObjectiveType.Percent, target: "1184.8", minBound: "-5", maxBound: "5" },
    ],
  },

  // --- Advanced: multi-sheet capital project finance ---
  {
    id: "capital-project-finance",
    name: "Capital Project Finance (4-month window)",
    description:
      "Multi-sheet financial model: adjust operating and capital expenditures across 4 months to hit cash flow and coverage targets.",
    complexity: "advanced",
    variables: [
      { name: "Fixed Opex Mo 1", type: InputType.Float, min: 50000, max: 200000 },
      { name: "Fixed Opex Mo 2", type: InputType.Float, min: 50000, max: 200000 },
      { name: "Fixed Opex Mo 3", type: InputType.Float, min: 50000, max: 200000 },
      { name: "Fixed Opex Mo 4", type: InputType.Float, min: 50000, max: 200000 },
      { name: "Capex Mo 1", type: InputType.Float, min: 0, max: 500000 },
      { name: "Capex Mo 2", type: InputType.Float, min: 0, max: 500000 },
      { name: "Capex Mo 3", type: InputType.Float, min: 0, max: 500000 },
      { name: "Capex Mo 4", type: InputType.Float, min: 0, max: 500000 },
    ],
    outcomeNames: [
      "NOI Mo 1", "NOI Mo 2", "NOI Mo 3", "NOI Mo 4",
      "DSCR Mo 1", "DSCR Mo 2", "DSCR Mo 3", "DSCR Mo 4",
      "Free CF Mo 1", "Free CF Mo 2", "Free CF Mo 3", "Free CF Mo 4",
      "Total NOI", "Min DSCR", "Total Free CF", "Total Spend",
    ],
    setup: {
      sheets: [
        // ── Assumptions ──
        {
          name: "Assumptions",
          merges: ["A1:E1"],
          boldRows: [0, 2, 5, 8],
          columnWidths: [26, 14, 14, 14, 14],
          data: [
            ["Capital Project - Financial Assumptions", null, null, null, null],
            [null, null, null, null, null],
            ["Revenue Projections",   "Mo 1",   "Mo 2",   "Mo 3",   "Mo 4"],
            ["Base Revenue ($)",      300000,    350000,   400000,   450000],
            [null, null, null, null, null],
            ["Financing", null, null, null, null],
            ["Monthly Debt Service ($)", 85000, null, null, null],
            [null, null, null, null, null],
            ["Cost Parameters", null, null, null, null],
            ["Tax Rate",              0.25, null, null, null],
            ["Maintenance Reserve %", 0.05, null, null, null],
          ],
        },
        // ── Model (Monthly P&L and Cash Flow) ──
        {
          name: "Model",
          merges: ["A1:E1"],
          boldRows: [0, 2],
          columnWidths: [22, 14, 14, 14, 14],
          inputHighlights: [
            "B5", "C5", "D5", "E5",   // Fixed Opex
            "B11", "C11", "D11", "E11", // Capex
          ],
          data: [
            ["Monthly Financial Model", null, null, null, null],
            [null, null, null, null, null],
            // Row 3 (Excel row 3): headers
            ["", "Mo 1", "Mo 2", "Mo 3", "Mo 4"],
            // Row 4: Revenue (from Assumptions row 4)
            ["Revenue",        "=Assumptions!B4", "=Assumptions!C4", "=Assumptions!D4", "=Assumptions!E4"],
            // Row 5: Fixed Opex — INPUT cells
            ["Fixed Opex",     0, 0, 0, 0],
            // Row 6: Gross Profit
            ["Gross Profit",   "=B4-B5", "=C4-C5", "=D4-D5", "=E4-E5"],
            // Row 7: Tax
            ["Tax",            "=B6*Assumptions!B10", "=C6*Assumptions!B10", "=D6*Assumptions!B10", "=E6*Assumptions!B10"],
            // Row 8: NOI — OUTPUT
            ["NOI",            "=B6-B7", "=C6-C7", "=D6-D7", "=E6-E7"],
            // Row 9: Maintenance Reserve
            ["Maint Reserve",  "=B4*Assumptions!B11", "=C4*Assumptions!B11", "=D4*Assumptions!B11", "=E4*Assumptions!B11"],
            // Row 10: Debt Service
            ["Debt Service",   "=Assumptions!B7", "=Assumptions!B7", "=Assumptions!B7", "=Assumptions!B7"],
            // Row 11: Capex — INPUT cells
            ["Capex",          0, 0, 0, 0],
            // Row 12: SNCF (Surplus Net Cash Flow)
            ["SNCF",           "=B8-B9-B10", "=C8-C9-C10", "=D8-D9-D10", "=E8-E9-E10"],
            // Row 13: DSCR — OUTPUT
            ["DSCR",           "=B8/B10", "=C8/C10", "=D8/D10", "=E8/E10"],
            // Row 14: Free Cash Flow — OUTPUT
            ["Free Cash Flow", "=B12-B11", "=C12-C11", "=D12-D11", "=E12-E11"],
          ],
        },
        // ── Summary ──
        {
          name: "Summary",
          merges: ["A1:B1"],
          boldRows: [0, 2],
          columnWidths: [22, 18],
          data: [
            ["Summary Metrics", null],
            [null, null],
            ["Metric", "Value"],
            // Row 4 (Excel row 4) — OUTPUT
            ["Total NOI",        "=SUM(Model!B8:E8)"],
            // Row 5 — OUTPUT
            ["Min Monthly DSCR", "=MIN(Model!B13:E13)"],
            // Row 6 — OUTPUT
            ["Total Free CF",    "=SUM(Model!B14:E14)"],
            // Row 7 — OUTPUT
            ["Total Spend",      "=SUM(Model!B5:E5)+SUM(Model!B11:E11)"],
          ],
        },
      ],
      // Input cell addresses (one per variable, in variable order)
      inputCells: [
        "Model!B5", "Model!C5", "Model!D5", "Model!E5",   // Fixed Opex Mo 1-4
        "Model!B11", "Model!C11", "Model!D11", "Model!E11", // Capex Mo 1-4
      ],
      // Output cell addresses (one per outcome, in outcomeNames order)
      outputCells: [
        "Model!B8",  "Model!C8",  "Model!D8",  "Model!E8",   // NOI Mo 1-4
        "Model!B13", "Model!C13", "Model!D13", "Model!E13",   // DSCR Mo 1-4
        "Model!B14", "Model!C14", "Model!D14", "Model!E14",   // Free CF Mo 1-4
        "Summary!B4", "Summary!B5", "Summary!B6", "Summary!B7", // Summary metrics
      ],
    },
    // Objectives: truth case Opex=[100k,110k,115k,120k], Capex=[200k,150k,100k,50k]
    // NOI: 150000, 180000, 213750, 247500
    // DSCR: 1.7647, 2.1176, 2.5147, 2.9118
    // Free CF: -150000, -72500, 8750, 90000
    // Total NOI: 791250, Min DSCR: 1.7647, Total Free CF: -123750, Total Spend: 945000
    objectives: [
      { type: ObjectiveType.Percent, target: "150000",  minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "180000",  minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "213750",  minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "247500",  minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "1.7647",  minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "2.1176",  minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "2.5147",  minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "2.9118",  minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "-150000", minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "-72500",  minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "8750",    minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "90000",   minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "791250",  minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "1.7647",  minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "-123750", minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "945000",  minBound: "-10", maxBound: "10" },
    ],
  },
  // --- Intermediate: annual budget allocation ---
  {
    id: "annual-budget-allocation",
    name: "Annual Budget Allocation",
    description:
      "Distribute a fixed annual budget across spending categories to simultaneously hit revenue, margin, and efficiency targets — the kind of planning most finance teams already do in Excel.",
    complexity: "intermediate",
    variables: [
      { name: "Headcount",       type: InputType.Float, min: 1000000, max: 5000000 },
      { name: "Software & Tools", type: InputType.Float, min: 50000,   max: 800000  },
      { name: "Marketing",        type: InputType.Float, min: 100000,  max: 2000000 },
      { name: "Training & Dev",   type: InputType.Float, min: 20000,   max: 500000  },
      { name: "Travel & Events",  type: InputType.Float, min: 10000,   max: 400000  },
      { name: "Contractors",      type: InputType.Float, min: 50000,   max: 1500000 },
    ],
    outcomeNames: [
      // Revenue drivers (4)
      "HC Revenue Contribution",
      "Marketing Revenue",
      "Contractor Output",
      "Events Pipeline Value",
      // Per-category ROI (4)
      "Revenue per HC $",
      "Marketing ROI",
      "Contractor ROI",
      "Events ROI",
      // Operational indices (5)
      "Software Productivity Index",
      "Training Impact Score",
      "Overhead Ratio",
      "Budget Utilization %",
      "Revenue to Cost Ratio",
      // Summary P&L (3)
      "Total Projected Revenue",
      "Gross Profit ($)",
      "Gross Profit Margin %",
    ],
    setup: {
      sheets: [
        // ── Assumptions ──
        {
          name: "Assumptions",
          merges: ["A1:B1"],
          boldRows: [0, 2, 7, 11],
          columnWidths: [26, 14],
          data: [
            ["Annual Budget Planning - Assumptions", null],
            [null, null],
            ["Revenue Multipliers", null],
            ["Headcount Revenue Multiplier",  3.50],
            ["Marketing Revenue Multiplier",  2.80],
            ["Contractor Revenue Multiplier", 1.20],
            ["Travel/Events Pipeline Rate",   4.50],
            [null, null],
            ["Productivity Factors", null],
            ["Software Productivity Rate",    1.80],
            ["Training Effectiveness Rate",   2.20],
            [null, null],
            ["Budget & Benchmarks", null],
            ["Total Annual Budget ($)",       4250000],
            ["Target Overhead Ratio",         0.30],
            ["Base Headcount",                85],
          ],
        },
        // ── Model ──
        {
          name: "Model",
          merges: ["A1:B1"],
          boldRows: [0, 2, 10, 17, 23],
          columnWidths: [26, 18],
          inputHighlights: ["B4", "B5", "B6", "B7", "B8", "B9"],
          data: [
            ["Annual Budget Model", null],
            [null, null],
            // Budget Allocations section (inputs in B4:B9)
            ["Budget Allocations", "Amount ($)"],
            ["Headcount",          0],
            ["Software & Tools",   0],
            ["Marketing",          0],
            ["Training & Dev",     0],
            ["Travel & Events",    0],
            ["Contractors",        0],
            ["Total Spend",        "=SUM(B4:B9)"],
            [null, null],
            // Revenue Drivers section
            ["Revenue Drivers",             "Value ($)"],
            ["HC Revenue Contribution",     "=B4*Assumptions!B4"],
            ["Marketing Revenue",           "=B6*Assumptions!B5"],
            ["Contractor Output",           "=B9*Assumptions!B6"],
            ["Events Pipeline Value",       "=B8*Assumptions!B7"],
            ["Total Projected Revenue",     "=SUM(B13:B16)"],
            [null, null],
            // Per-Category ROI section
            ["ROI Metrics",                 "Ratio"],
            ["Revenue per HC $",            "=B13/B4"],
            ["Marketing ROI",               "=B14/B6"],
            ["Contractor ROI",              "=B15/B9"],
            ["Events ROI",                  "=B16/B8"],
            [null, null],
            // Operational Indices section
            ["Operational Indices",         "Value"],
            ["Software Productivity Index", "=B5*Assumptions!B10/Assumptions!B14*100"],
            ["Training Impact Score",       "=B7*Assumptions!B11/Assumptions!B14*100"],
            ["Overhead Ratio",              "=(B5+B7+B8)/B17"],
            ["Budget Utilization %",        "=B10/Assumptions!B14"],
            ["Revenue to Cost Ratio",       "=B17/B10"],
          ],
        },
        // ── Summary ──
        {
          name: "Summary",
          merges: ["A1:B1"],
          boldRows: [0, 2],
          columnWidths: [26, 18],
          data: [
            ["Budget Summary", null],
            [null, null],
            ["KPI",                     "Value"],
            ["Total Projected Revenue", "=Model!B17"],
            ["Total Operating Cost",    "=Model!B10"],
            ["Gross Profit ($)",        "=Model!B17-Model!B10"],
            ["Gross Profit Margin %",   "=(Model!B17-Model!B10)/Model!B17"],
          ],
        },
      ],
      // Input cell addresses (one per variable, in variable order)
      inputCells: [
        "Model!B4",  // Headcount
        "Model!B5",  // Software & Tools
        "Model!B6",  // Marketing
        "Model!B7",  // Training & Dev
        "Model!B8",  // Travel & Events
        "Model!B9",  // Contractors
      ],
      // Output cell addresses (one per outcome, in outcomeNames order)
      outputCells: [
        // Revenue drivers
        "Model!B13", "Model!B14", "Model!B15", "Model!B16",
        // Per-category ROI
        "Model!B20", "Model!B21", "Model!B22", "Model!B23",
        // Operational indices
        "Model!B26", "Model!B27", "Model!B28", "Model!B29", "Model!B30",
        // Summary P&L
        "Summary!B4", "Summary!B6", "Summary!B7",
      ],
    },
    // Truth case: HC=2.5M, SW=350K, Mktg=750K, Training=180K, Travel=120K, Contractors=350K
    // Total spend: 4,250,000
    // HC Rev: 8,750,000  |  Mktg Rev: 2,100,000  |  Contractor: 420,000  |  Events: 540,000
    // Total Rev: 11,810,000  |  Gross Profit: 7,560,000  |  Margin: 64.0%
    objectives: [
      // Revenue drivers
      { type: ObjectiveType.Percent, target: "8750000",  minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "2100000",  minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "420000",   minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "540000",   minBound: "-10", maxBound: "10" },
      // ROI metrics
      { type: ObjectiveType.Percent, target: "3.5",      minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "2.8",      minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "1.2",      minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "4.5",      minBound: "-10", maxBound: "10" },
      // Operational indices
      { type: ObjectiveType.Percent, target: "14.82",    minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "9.32",     minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "0.055",    minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "1.0",      minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "2.78",     minBound: "-10", maxBound: "10" },
      // Summary P&L
      { type: ObjectiveType.Percent, target: "11810000", minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "7560000",  minBound: "-10", maxBound: "10" },
      { type: ObjectiveType.Percent, target: "0.640",    minBound: "-10", maxBound: "10" },
    ],
  },
];

export function getExampleById(id: string): Example | undefined {
  return EXAMPLES.find((e) => e.id === id);
}
