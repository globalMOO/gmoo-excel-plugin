import React, { useState } from "react";
import {
  makeStyles,
  tokens,
  Button,
  Input,
  Text,
  Dropdown,
  Option,
  Spinner,
  MessageBar,
  MessageBarBody,
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
} from "@fluentui/react-components";
import { CursorClick20Regular } from "@fluentui/react-icons";
import type { GmooClient } from "../services/gmooApi";
import type { EvalConfig } from "../services/excelService";
import { evaluateCase, readSelectedRangeWithValues } from "../services/excelService";
import { ObjectiveType } from "../types/gmoo";
import type { InputVariable } from "../types/workbookState";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    padding: "16px",
  },
  smallInput: {
    width: "80px",
  },
  typeDropdown: {
    minWidth: "110px",
  },
  buttonRow: {
    display: "flex",
    justifyContent: "space-between",
    marginTop: "12px",
  },
});

const OBJECTIVE_TYPE_OPTIONS = [
  { value: ObjectiveType.Value, label: "Value" },
  { value: ObjectiveType.Percent, label: "Percent" },
  { value: ObjectiveType.LessThan, label: "Less Than" },
  { value: ObjectiveType.LessThanEqual, label: "Less Than Equal" },
  { value: ObjectiveType.GreaterThan, label: "Greater Than" },
  { value: ObjectiveType.GreaterThanEqual, label: "Greater Than Equal" },
  { value: ObjectiveType.Minimize, label: "Minimize" },
  { value: ObjectiveType.Maximize, label: "Maximize" },
];

const NO_TARGET_TYPES = new Set([ObjectiveType.Minimize, ObjectiveType.Maximize]);

/** Parse a string cell into an ObjectiveType, or null if unrecognized. */
function parseObjectiveType(raw: unknown): ObjectiveType | null {
  if (raw === null || raw === undefined) return null;
  const s = String(raw).trim().toLowerCase();
  if (!s) return null;
  if (s === "value" || s === "v" || s === "equal" || s === "eq" || s === "=")
    return ObjectiveType.Value;
  if (s === "percent" || s === "pct" || s === "%") return ObjectiveType.Percent;
  if (s === "lessthan" || s === "less than" || s === "lt" || s === "<")
    return ObjectiveType.LessThan;
  if (s === "lessthanequal" || s === "less than equal" || s === "lte" || s === "<=")
    return ObjectiveType.LessThanEqual;
  if (s === "greaterthan" || s === "greater than" || s === "gt" || s === ">")
    return ObjectiveType.GreaterThan;
  if (s === "greaterthanequal" || s === "greater than equal" || s === "gte" || s === ">=")
    return ObjectiveType.GreaterThanEqual;
  if (s === "minimize" || s === "min") return ObjectiveType.Minimize;
  if (s === "maximize" || s === "max") return ObjectiveType.Maximize;
  return null;
}

function cellToString(raw: unknown): string {
  if (raw === null || raw === undefined) return "";
  return String(raw);
}

export interface ObjectiveRowData {
  type: ObjectiveType;
  target: string;
  minBound: string;
  maxBound: string;
}

interface ObjectiveSetupProps {
  client: GmooClient | null;
  trialId: number;
  outcomeNames: string[];
  inputCases: number[][];
  /** Used to synthesize an initial input (per-variable midpoint) when
   *  inputCases is empty — happens on Resume-from-Trial because the project
   *  DTO usually omits the training-cases array. */
  variables: InputVariable[];
  evalConfig: EvalConfig | null;
  initialObjectives?: ObjectiveRowData[];
  /** Default objectives from a loaded example */
  exampleObjectives?: ObjectiveRowData[];
  onComplete: (objectiveId: number, objectiveRows: ObjectiveRowData[]) => void;
  onBack: () => void;
  /** Optional slot rendered above the body (e.g. PickExistingBar). */
  headerSlot?: React.ReactNode;
}

function makeBlankRows(outcomeNames: string[]): ObjectiveRowData[] {
  return outcomeNames.map(() => ({
    type: ObjectiveType.Percent,
    target: "0",
    minBound: "-5",
    maxBound: "5",
  }));
}

export const ObjectiveSetup: React.FC<ObjectiveSetupProps> = ({
  client,
  trialId,
  outcomeNames,
  inputCases,
  variables,
  evalConfig,
  initialObjectives,
  exampleObjectives,
  onComplete,
  onBack,
  headerSlot,
}) => {
  const styles = useStyles();

  const [objectives, setObjectives] = useState<ObjectiveRowData[]>(
    initialObjectives ?? exampleObjectives ?? makeBlankRows(outcomeNames)
  );
  const [initialCaseIndex, setInitialCaseIndex] = useState(0);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const updateObjective = (index: number, field: keyof ObjectiveRowData, value: string) => {
    setObjectives(
      objectives.map((obj, i) =>
        i === index ? { ...obj, [field]: value } : obj
      )
    );
  };

  /**
   * Populate the objectives table from the user's current Excel selection.
   *
   * Column layouts, inferred from the shape of the range:
   *   1 col  → Target
   *   2 cols → Type | Target   (if col-1 parses as a type for any row)
   *            else Target | MinBound
   *   3 cols → Type | Target | (symmetric ± Bound)  if col-1 is a type
   *            else Target | MinBound | MaxBound
   *   4 cols → Type | Target | MinBound | MaxBound
   *
   * Row count must match `outcomeNames.length`. Missing/unparseable fields
   * leave the existing value untouched.
   */
  const loadObjectivesFromSelection = async () => {
    setError(null);
    try {
      const { values } = await readSelectedRangeWithValues();
      // Trim fully-blank trailing rows
      const rows = values.filter(
        (r) => r && r.some((c) => c !== null && c !== undefined && String(c).trim() !== "")
      );
      if (rows.length === 0) {
        setError("Selected range is empty.");
        return;
      }
      if (rows.length !== outcomeNames.length) {
        setError(
          `Selection has ${rows.length} row(s) but there are ${outcomeNames.length} outcome(s). Select exactly one row per outcome.`
        );
        return;
      }

      const nCols = Math.max(...rows.map((r) => r.length));
      // Does the first column parse as a type for at least one row?
      const firstColIsType = rows.some((r) => parseObjectiveType(r[0]) !== null);

      const next = objectives.map((obj, i) => {
        const row = rows[i];
        const updated: ObjectiveRowData = { ...obj };

        if (nCols === 1) {
          updated.target = cellToString(row[0]) || obj.target;
        } else if (firstColIsType) {
          const t = parseObjectiveType(row[0]);
          if (t) updated.type = t;
          if (nCols >= 2) updated.target = cellToString(row[1]) || obj.target;
          if (nCols === 3) {
            // Single "± bound" column → apply to both min and max (symmetric)
            const b = cellToString(row[2]);
            if (b) {
              const absB = String(Math.abs(parseFloat(b) || 0));
              updated.minBound = `-${absB}`;
              updated.maxBound = absB;
            }
          } else if (nCols >= 4) {
            if (cellToString(row[2])) updated.minBound = cellToString(row[2]);
            if (cellToString(row[3])) updated.maxBound = cellToString(row[3]);
          }
        } else {
          // No type column — columns are Target [, Min [, Max]]
          updated.target = cellToString(row[0]) || obj.target;
          if (nCols >= 2 && cellToString(row[1])) updated.minBound = cellToString(row[1]);
          if (nCols >= 3 && cellToString(row[2])) updated.maxBound = cellToString(row[2]);
        }
        return updated;
      });
      setObjectives(next);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to read selected range.");
    }
  };

  const handleSubmit = async () => {
    if (!client) return;

    setIsSubmitting(true);
    setError(null);

    try {
      // Prefer a real training case. Fall back to per-variable midpoints when
      // Resume-from-Trial gave us no inputCases (project DTO omits them) so
      // the user isn't stuck — the value only seeds the first inverse
      // iteration; the optimizer takes over from there.
      const initialInput =
        inputCases[initialCaseIndex] ??
        inputCases[0] ??
        (variables.length > 0
          ? variables.map((v) => (v.min + v.max) / 2)
          : null);
      if (!initialInput) {
        setError(
          "No initial input could be derived. Go back and re-define the model's variables."
        );
        setIsSubmitting(false);
        return;
      }
      const targetValues = objectives.map((o) => parseFloat(o.target) || 0);

      // Evaluate the initial input case through Excel to get real outputs
      // (using target values here would make the API think the objective is already satisfied)
      let initialOutput: number[];
      if (evalConfig) {
        const evalResult = await evaluateCase(evalConfig, initialInput);
        initialOutput = evalResult.outputs;
      } else {
        // Fallback: zeros (not target values — that causes the API to mark objective as satisfied)
        initialOutput = new Array(outcomeNames.length).fill(0);
      }

      const result = await client.loadObjectives(
        trialId,
        targetValues,
        objectives.map((o) => o.type),
        initialInput,
        initialOutput,
        0,
        objectives.map((o) => parseFloat(o.minBound) || 0),
        objectives.map((o) => parseFloat(o.maxBound) || 0)
      );

      onComplete(result.id, objectives);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to set objectives.");
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <div className={styles.container}>
      {headerSlot}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: "8px" }}>
        <Text weight="semibold" size={400}>
          Set Objectives
        </Text>
        <Button
          icon={<CursorClick20Regular />}
          size="small"
          appearance="subtle"
          onClick={loadObjectivesFromSelection}
          title="Populate Type / Target / Min / Max from the currently-selected cells (1–4 columns; any sheet or workbook). One row per outcome."
        >
          Load from selection
        </Button>
      </div>
      <Text size={200}>
        Define optimization targets for each outcome.
      </Text>

      <Table size="extra-small">
        <TableHeader>
          <TableRow>
            <TableHeaderCell>Outcome</TableHeaderCell>
            <TableHeaderCell>Type</TableHeaderCell>
            <TableHeaderCell>Target</TableHeaderCell>
            <TableHeaderCell>Min Bound</TableHeaderCell>
            <TableHeaderCell>Max Bound</TableHeaderCell>
          </TableRow>
        </TableHeader>
        <TableBody>
          {outcomeNames.map((name, i) => (
            <TableRow key={i}>
              <TableCell>
                <Text size={200}>{name}</Text>
              </TableCell>
              <TableCell>
                <Dropdown
                  className={styles.typeDropdown}
                  size="small"
                  value={objectives[i].type}
                  selectedOptions={[objectives[i].type]}
                  onOptionSelect={(_, data) =>
                    updateObjective(i, "type", (data.optionValue as ObjectiveType) ?? ObjectiveType.Value)
                  }
                >
                  {OBJECTIVE_TYPE_OPTIONS.map((opt) => (
                    <Option key={opt.value} value={opt.value}>
                      {opt.label}
                    </Option>
                  ))}
                </Dropdown>
              </TableCell>
              <TableCell>
                <Input
                  className={styles.smallInput}
                  size="small"
                  value={objectives[i].target}
                  onChange={(_, data) =>
                    updateObjective(i, "target", data.value)
                  }
                  disabled={NO_TARGET_TYPES.has(objectives[i].type)}
                />
              </TableCell>
              <TableCell>
                <Input
                  className={styles.smallInput}
                  size="small"
                  value={objectives[i].minBound}
                  onChange={(_, data) =>
                    updateObjective(i, "minBound", data.value)
                  }
                />
              </TableCell>
              <TableCell>
                <Input
                  className={styles.smallInput}
                  size="small"
                  value={objectives[i].maxBound}
                  onChange={(_, data) =>
                    updateObjective(i, "maxBound", data.value)
                  }
                />
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>

      <div style={{ display: "flex", gap: "12px", alignItems: "center" }}>
        <Text size={200}>Initial case index:</Text>
        <Input
          size="small"
          type="number"
          value={String(initialCaseIndex)}
          onChange={(_, data) => setInitialCaseIndex(parseInt(data.value) || 0)}
          style={{ width: "80px" }}
        />
        <Text size={100}>(0-based, default first case)</Text>
      </div>

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
        Submitting creates a new optimization run against this trial. Existing
        objectives in this trial aren't modified — you can switch back to them
        from the dropdown above.
      </Text>

      <div className={styles.buttonRow}>
        <Button appearance="secondary" onClick={onBack}>
          Back
        </Button>
        <Button appearance="primary" onClick={handleSubmit} disabled={isSubmitting}>
          {isSubmitting ? <Spinner size="tiny" /> : "Submit Objectives & Start Optimization"}
        </Button>
      </div>
    </div>
  );
};
