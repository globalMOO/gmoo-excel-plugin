import React, { useState, useEffect } from "react";
import {
  makeStyles,
  tokens,
  Button,
  Text,
  Spinner,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  ProgressBar,
  RadioGroup,
  Radio,
  Card,
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  Badge,
} from "@fluentui/react-components";
import { CursorClick20Regular } from "@fluentui/react-icons";
import type { GmooClient } from "../services/gmooApi";
import type { InputVariable } from "../types/workbookState";
import type { EvalConfig } from "../services/excelService";
import { createTemplateSheet, evaluateAllCases, readSelectedRange, loadStateSheet, saveStateSheet, type VsmeStateData } from "../services/excelService";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    padding: "16px",
  },
  progressSection: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
  buttonRow: {
    display: "flex",
    justifyContent: "space-between",
    marginTop: "12px",
  },
  cellPickRow: {
    display: "flex",
    gap: "4px",
    alignItems: "center",
  },
  cellAddress: {
    fontFamily: "monospace",
    fontSize: "11px",
    minWidth: "120px",
  },
});

interface CaseEvaluationProps {
  client: GmooClient | null;
  modelName: string;
  projectId: number;
  variables: InputVariable[];
  outcomeNames: string[];
  inputCases: number[][];
  /** Optional formulas from a loaded example, keyed by outcome name */
  formulas?: Record<string, string>;
  /** Pre-built eval config from example auto-setup (sheet already exists) */
  initialEvalConfig?: EvalConfig;
  onComplete: (trialId: number, evalConfig: EvalConfig) => void;
  onBack: () => void;
}

export const CaseEvaluation: React.FC<CaseEvaluationProps> = ({
  client,
  modelName,
  projectId,
  variables,
  outcomeNames,
  inputCases,
  formulas,
  initialEvalConfig,
  onComplete,
  onBack,
}) => {
  const styles = useStyles();

  const [mode, setMode] = useState<"template" | "existing">(
    initialEvalConfig?.inputCells ? "existing" : "template"
  );
  const [evalConfig, setEvalConfig] = useState<EvalConfig | null>(initialEvalConfig ?? null);
  const [isCreatingSheet, setIsCreatingSheet] = useState(false);
  const [isEvaluating, setIsEvaluating] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0 });
  const [error, setError] = useState<string | null>(null);
  const [sheetCreated, setSheetCreated] = useState(!!initialEvalConfig);

  // Per-variable cell mapping for existing sheet mode
  const [inputCellMap, setInputCellMap] = useState<string[]>(
    () => initialEvalConfig?.inputCells ?? new Array(variables.length).fill("")
  );
  const [outputCellMap, setOutputCellMap] = useState<string[]>(
    () => initialEvalConfig?.outputCells ?? new Array(outcomeNames.length).fill("")
  );
  const [isPicking, setIsPicking] = useState<string | null>(null); // "input-0", "output-2", etc.
  const [stateLoaded, setStateLoaded] = useState(false);
  const [isLoadingState, setIsLoadingState] = useState(false);

  // Auto-detect _VSME_State sheet when switching to "existing" mode
  useEffect(() => {
    if (mode !== "existing" || sheetCreated || stateLoaded || initialEvalConfig) return;

    let cancelled = false;
    setIsLoadingState(true);

    loadStateSheet().then((data: VsmeStateData | null) => {
      if (cancelled || !data) {
        setIsLoadingState(false);
        return;
      }

      // Match by variable name to fill input cells
      const newInputMap = new Array(variables.length).fill("");
      for (let i = 0; i < variables.length; i++) {
        const match = data.variables.find((v) => v.name === variables[i].name);
        if (match?.inputCell) newInputMap[i] = match.inputCell;
      }

      // Match by outcome name to fill output cells
      const newOutputMap = new Array(outcomeNames.length).fill("");
      for (let i = 0; i < outcomeNames.length; i++) {
        const match = data.outcomes.find((o) => o.name === outcomeNames[i]);
        if (match?.outputCell) newOutputMap[i] = match.outputCell;
      }

      setInputCellMap(newInputMap);
      setOutputCellMap(newOutputMap);
      setStateLoaded(true);
      setIsLoadingState(false);
    }).catch(() => {
      setIsLoadingState(false);
    });

    return () => { cancelled = true; };
  }, [mode, sheetCreated, stateLoaded, initialEvalConfig, variables, outcomeNames]);

  const handleCreateTemplate = async () => {
    setIsCreatingSheet(true);
    setError(null);
    try {
      const sheetName = `${modelName.substring(0, 20)} Model Def`;
      const config = await createTemplateSheet({
        modelName,
        variables,
        outcomeNames,
        sheetName,
        formulas,
      });
      setEvalConfig(config);
      setSheetCreated(true);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to create template sheet.");
    } finally {
      setIsCreatingSheet(false);
    }
  };

  const handlePickCell = async (type: "input" | "output", index: number) => {
    const key = `${type}-${index}`;
    setIsPicking(key);
    try {
      const address = await readSelectedRange();
      if (type === "input") {
        setInputCellMap((prev) => {
          const next = [...prev];
          next[index] = address;
          return next;
        });
      } else {
        setOutputCellMap((prev) => {
          const next = [...prev];
          next[index] = address;
          return next;
        });
      }
    } catch {
      // cancelled
    } finally {
      setIsPicking(null);
    }
  };

  const handleConfirmExisting = () => {
    const missingInputs = inputCellMap.filter((c) => !c.trim());
    const missingOutputs = outputCellMap.filter((c) => !c.trim());
    if (missingInputs.length > 0 || missingOutputs.length > 0) {
      setError(
        `All cells must be mapped. Missing: ${missingInputs.length} input(s), ${missingOutputs.length} output(s).`
      );
      return;
    }

    const config: EvalConfig = {
      variableCount: variables.length,
      outcomeCount: outcomeNames.length,
      inputCells: inputCellMap.map((c) => c.trim()),
      outputCells: outputCellMap.map((c) => c.trim()),
    };
    setEvalConfig(config);
    setSheetCreated(true);
    setError(null);
  };

  const handleEvaluate = async () => {
    if (!evalConfig || !client) return;

    setIsEvaluating(true);
    setError(null);

    try {
      const { outputCases, errors } = await evaluateAllCases(
        evalConfig,
        inputCases,
        (current, total) => setProgress({ current, total })
      );

      if (errors.length > 0) {
        setError(`Formula errors detected:\n${errors.join("\n")}`);
        return;
      }

      // Submit output cases to API
      const trial = await client.loadOutputCases(
        projectId,
        outcomeNames.length,
        outputCases
      );

      // Save cell mappings to _VSME_State sheet for future re-use
      const inputCells = evalConfig.inputCells ?? variables.map((_, i) => {
        const col = String.fromCharCode(64 + (evalConfig.inputStartCol ?? 2) + i);
        return `${evalConfig.sheetName}!${col}${evalConfig.inputStartRow ?? 7}`;
      });
      const outputCells = evalConfig.outputCells ?? outcomeNames.map((_, i) => {
        const col = String.fromCharCode(64 + (evalConfig.outputStartCol ?? 2));
        return `${evalConfig.sheetName}!${col}${(evalConfig.outputStartRow ?? 11) + i}`;
      });
      try {
        await saveStateSheet({
          variables: variables.map((v, i) => ({
            name: v.name,
            type: v.type,
            min: v.min,
            max: v.max,
            inputCell: inputCells[i],
          })),
          outcomes: outcomeNames.map((name, i) => ({
            name,
            outputCell: outputCells[i],
          })),
        });
      } catch {
        // Non-critical — don't fail the evaluation if state save fails
        console.warn("[VSME] Failed to save state sheet");
      }

      onComplete(trial.id, evalConfig);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Evaluation failed.");
    } finally {
      setIsEvaluating(false);
    }
  };

  const allInputsMapped = inputCellMap.every((c) => c.trim());
  const allOutputsMapped = outputCellMap.every((c) => c.trim());

  return (
    <div className={styles.container}>
      <Text weight="semibold" size={400}>
        Evaluate Input Cases
      </Text>
      <Text size={200}>
        The API generated {inputCases.length} input cases. Your Excel formulas will compute
        the outputs for each case to train the model.
      </Text>

      <RadioGroup
        value={mode}
        onChange={(_, data) => {
          setMode(data.value as "template" | "existing");
          setSheetCreated(false);
          setEvalConfig(null);
          setError(null);
        }}
      >
        <Radio value="template" label="Create Template Sheet (recommended)" />
        <Radio value="existing" label="Use Existing Sheet" />
      </RadioGroup>

      {mode === "template" && !sheetCreated && (
        <Button
          appearance="primary"
          onClick={handleCreateTemplate}
          disabled={isCreatingSheet}
        >
          {isCreatingSheet ? <Spinner size="tiny" /> : "Create Template Sheet"}
        </Button>
      )}

      {mode === "template" && sheetCreated && (
        <MessageBar intent="success">
          <MessageBarBody>
            <MessageBarTitle>Template Created</MessageBarTitle>
            Fill in your outcome formulas in column B (rows 11+), referencing the Current Value
            cells in row 7. Then click "Evaluate All Cases".
          </MessageBarBody>
        </MessageBar>
      )}

      {mode === "existing" && !sheetCreated && isLoadingState && (
        <Spinner label="Checking for saved configuration..." size="small" />
      )}

      {mode === "existing" && !sheetCreated && !isLoadingState && (
        <Card>
          <div style={{ padding: "12px", display: "flex", flexDirection: "column", gap: "12px" }}>
            {stateLoaded && (
              <MessageBar intent="success">
                <MessageBarBody>
                  Cell mappings loaded from previous run. Review and adjust if needed.
                </MessageBarBody>
              </MessageBar>
            )}
            <Text weight="semibold" size={300}>
              Map Input Variables
            </Text>
            <Text size={200}>
              For each input, select the cell where the add-in should write values.
              Cells can be on different sheets.
            </Text>
            <Table size="extra-small">
              <TableHeader>
                <TableRow>
                  <TableHeaderCell>Variable</TableHeaderCell>
                  <TableHeaderCell>Cell</TableHeaderCell>
                  <TableHeaderCell></TableHeaderCell>
                </TableRow>
              </TableHeader>
              <TableBody>
                {variables.map((v, i) => (
                  <TableRow key={`input-${i}`}>
                    <TableCell>
                      <Text size={200}>{v.name}</Text>
                    </TableCell>
                    <TableCell>
                      <Text size={200} className={styles.cellAddress}>
                        {inputCellMap[i] || (
                          <span style={{ color: tokens.colorNeutralForeground4 }}>not set</span>
                        )}
                      </Text>
                    </TableCell>
                    <TableCell>
                      <Button
                        icon={<CursorClick20Regular />}
                        size="small"
                        appearance="subtle"
                        onClick={() => handlePickCell("input", i)}
                        disabled={isPicking !== null}
                      >
                        {isPicking === `input-${i}` ? "Click cell..." : "Pick"}
                      </Button>
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
            {allInputsMapped && (
              <Badge appearance="filled" color="success" style={{ alignSelf: "flex-start" }}>
                All inputs mapped
              </Badge>
            )}

            <Text weight="semibold" size={300} style={{ marginTop: "8px" }}>
              Map Outcomes
            </Text>
            <Text size={200}>
              For each outcome, select the cell containing the formula result.
            </Text>
            <Table size="extra-small">
              <TableHeader>
                <TableRow>
                  <TableHeaderCell>Outcome</TableHeaderCell>
                  <TableHeaderCell>Cell</TableHeaderCell>
                  <TableHeaderCell></TableHeaderCell>
                </TableRow>
              </TableHeader>
              <TableBody>
                {outcomeNames.map((name, i) => (
                  <TableRow key={`output-${i}`}>
                    <TableCell>
                      <Text size={200}>{name}</Text>
                    </TableCell>
                    <TableCell>
                      <Text size={200} className={styles.cellAddress}>
                        {outputCellMap[i] || (
                          <span style={{ color: tokens.colorNeutralForeground4 }}>not set</span>
                        )}
                      </Text>
                    </TableCell>
                    <TableCell>
                      <Button
                        icon={<CursorClick20Regular />}
                        size="small"
                        appearance="subtle"
                        onClick={() => handlePickCell("output", i)}
                        disabled={isPicking !== null}
                      >
                        {isPicking === `output-${i}` ? "Click cell..." : "Pick"}
                      </Button>
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
            {allOutputsMapped && (
              <Badge appearance="filled" color="success" style={{ alignSelf: "flex-start" }}>
                All outcomes mapped
              </Badge>
            )}

            <Button
              appearance="primary"
              onClick={handleConfirmExisting}
              disabled={!allInputsMapped || !allOutputsMapped}
              style={{ marginTop: "8px" }}
            >
              Confirm Cell Mapping
            </Button>
          </div>
        </Card>
      )}

      {sheetCreated && (
        <Button
          appearance="primary"
          onClick={handleEvaluate}
          disabled={isEvaluating}
        >
          {isEvaluating ? <Spinner size="tiny" /> : "Evaluate All Cases"}
        </Button>
      )}

      {isEvaluating && (
        <div className={styles.progressSection}>
          <Text size={200}>
            Evaluating case {progress.current} of {progress.total}...
          </Text>
          <ProgressBar
            value={progress.total > 0 ? progress.current / progress.total : 0}
          />
        </div>
      )}

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>
            <MessageBarTitle>Error</MessageBarTitle>
            <pre style={{ whiteSpace: "pre-wrap", fontSize: "12px" }}>{error}</pre>
          </MessageBarBody>
        </MessageBar>
      )}

      <div className={styles.buttonRow}>
        <Button appearance="secondary" onClick={onBack} disabled={isEvaluating}>
          Back
        </Button>
      </div>
    </div>
  );
};
