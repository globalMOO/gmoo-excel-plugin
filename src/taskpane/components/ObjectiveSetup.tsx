import React, { useState } from "react";
import {
  makeStyles,
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
import type { GmooClient } from "../services/gmooApi";
import type { EvalConfig } from "../services/excelService";
import { evaluateCase } from "../services/excelService";
import { ObjectiveType } from "../types/gmoo";

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
  evalConfig: EvalConfig | null;
  initialObjectives?: ObjectiveRowData[];
  /** Default objectives from a loaded example */
  exampleObjectives?: ObjectiveRowData[];
  onComplete: (objectiveId: number, objectiveRows: ObjectiveRowData[]) => void;
  onBack: () => void;
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
  evalConfig,
  initialObjectives,
  exampleObjectives,
  onComplete,
  onBack,
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

  const handleSubmit = async () => {
    if (!client) return;

    setIsSubmitting(true);
    setError(null);

    try {
      const initialInput = inputCases[initialCaseIndex] ?? inputCases[0];
      const targetValues = objectives.map((o) => parseFloat(o.target) || 0);

      // Evaluate the initial input case through Excel to get real outputs
      // (using target values here would make the API think the objective is already satisfied)
      let initialOutput: number[];
      if (evalConfig) {
        const evalResult = await evaluateCase(evalConfig, initialInput);
        initialOutput = evalResult.outputs;
        console.log("[ObjectiveSetup] Evaluated initial case outputs:", initialOutput);
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
      <Text weight="semibold" size={400}>
        Set Objectives
      </Text>
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

      <div className={styles.buttonRow}>
        <Button appearance="secondary" onClick={onBack}>
          Back
        </Button>
        <Button appearance="primary" onClick={handleSubmit} disabled={isSubmitting}>
          {isSubmitting ? <Spinner size="tiny" /> : "Set Objectives & Continue"}
        </Button>
      </div>
    </div>
  );
};
