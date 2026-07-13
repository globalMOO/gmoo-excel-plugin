import React, { useState, useEffect } from "react";
import {
  makeStyles,
  tokens,
  Button,
  Input,
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
import { CursorClick20Regular, Add20Regular, Delete20Regular } from "@fluentui/react-icons";
import type { GmooClient } from "../services/gmooApi";
import type { InputVariable } from "../types/workbookState";
import type { CalcMode, EvalConfig } from "../services/excelService";
import { createTemplateSheet, evaluateAllCases, readSelectedRange, readSelectedRangeWithValues, readSelectedCellAddresses, loadStateSheet, saveStateSheet, type GmooStateData } from "../services/excelService";

/** Parse a 2-D values array into outcome names (first column, blank rows skipped). */
function parseOutcomesFromRange(values: unknown[][]): string[] {
  const names: string[] = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i] ?? [];
    const cell = row[0];
    if (cell === null || cell === undefined) continue;
    const s = String(cell).trim();
    if (!s) continue;
    names.push(s);
  }
  return names;
}

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
  /** Seed for the editable Outcomes list (example/resume) — empty for a fresh
   *  manual flow, where the user now defines outcomes here at the Evaluate step. */
  initialOutcomeNames: string[];
  inputCases: number[][];
  /** Optional formulas from a loaded example, keyed by outcome name */
  formulas?: Record<string, string>;
  /** Pre-built eval config from example auto-setup (sheet already exists) */
  initialEvalConfig?: EvalConfig;
  onComplete: (trialId: number, evalConfig: EvalConfig, outcomeNames: string[]) => void;
  onBack: () => void;
  /** Recover the project's DOE input cases (Resume-from-Project lands with none)
   *  so the user can evaluate and create a new trial. Resolves false when the
   *  API doesn't carry them. */
  onReloadInputCases?: () => Promise<boolean>;
  /** Optional slot rendered above the body (e.g. PickExistingBar). */
  headerSlot?: React.ReactNode;
}

export const CaseEvaluation: React.FC<CaseEvaluationProps> = ({
  client,
  modelName,
  projectId,
  variables,
  initialOutcomeNames,
  inputCases,
  formulas,
  initialEvalConfig,
  onComplete,
  onBack,
  onReloadInputCases,
  headerSlot,
}) => {
  const styles = useStyles();

  // Outcomes are defined here (not at Define Model) — the API only needs them
  // at training time (loadOutputCases). Seed from example/resume, else one blank.
  const [outcomeNames, setOutcomeNames] = useState<string[]>(
    () => (initialOutcomeNames.length > 0 ? initialOutcomeNames : [""])
  );

  const [mode, setMode] = useState<"template" | "existing">(
    initialEvalConfig?.inputCells ? "existing" : "template"
  );
  const [evalConfig, setEvalConfig] = useState<EvalConfig | null>(initialEvalConfig ?? null);
  // Recalculation mode for evaluations. "recalculate" (changed-only) is
  // correct for virtually all workbooks and far faster on large models;
  // "full" recomputes every formula each case (pre-option behavior).
  const [calcMode, setCalcMode] = useState<CalcMode>(
    initialEvalConfig?.calcMode ?? "recalculate"
  );
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
  // Recovery from the "no input cases" dead-end (Resume-from-Project).
  const [isReloadingCases, setIsReloadingCases] = useState(false);
  const [reloadError, setReloadError] = useState<string | null>(null);

  const handleReloadInputCases = async () => {
    if (!onReloadInputCases) return;
    setIsReloadingCases(true);
    setReloadError(null);
    try {
      const ok = await onReloadInputCases();
      if (!ok) {
        setReloadError(
          "This project's input cases aren't available from the API (the DOE is fixed when a project is created and isn't returned on reload). Use \"Switch trial\" above to pick an already-trained trial, or start a new model."
        );
      }
      // On success the parent updates inputCases; this view re-renders into the
      // normal evaluation flow automatically.
    } catch (err) {
      setReloadError(err instanceof Error ? err.message : "Failed to reload input cases.");
    } finally {
      setIsReloadingCases(false);
    }
  };

  // Auto-detect _GMOO_State sheet when switching to "existing" mode
  useEffect(() => {
    if (mode !== "existing" || sheetCreated || stateLoaded || initialEvalConfig) return;

    let cancelled = false;
    setIsLoadingState(true);

    loadStateSheet().then((data: GmooStateData | null) => {
      if (cancelled || !data) {
        setIsLoadingState(false);
        return;
      }
      // If the sheet was tagged with a project id, only auto-fill when it
      // matches — otherwise we'd silently apply another project's cells to
      // this one whenever variable names happen to overlap.
      if (typeof data.projectId === "number" && data.projectId !== projectId) {
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
      if (data.calcMode) setCalcMode(data.calcMode);
      setStateLoaded(true);
      setIsLoadingState(false);
    }).catch(() => {
      setIsLoadingState(false);
    });

    return () => { cancelled = true; };
  }, [mode, sheetCreated, stateLoaded, initialEvalConfig, variables, outcomeNames, projectId]);

  // Keep the existing-sheet output cell map sized to the editable outcome list,
  // preserving already-picked cells when the user adds/removes rows.
  useEffect(() => {
    setOutputCellMap((prev) => {
      if (prev.length === outcomeNames.length) return prev;
      const next = new Array(outcomeNames.length).fill("");
      for (let i = 0; i < Math.min(prev.length, next.length); i++) next[i] = prev[i];
      return next;
    });
  }, [outcomeNames.length]);

  // Blank outcome names auto-fill to "Outcome N" at commit time, mirroring the
  // old Define-Model behavior. The editor itself shows the raw (possibly blank)
  // names with placeholders.
  const filledOutcomeNames = outcomeNames.map((n, i) => n.trim() || `Outcome ${i + 1}`);

  const addOutcome = () => setOutcomeNames([...outcomeNames, ""]);
  const removeOutcome = (index: number) => {
    if (outcomeNames.length <= 1) return;
    setOutcomeNames(outcomeNames.filter((_, i) => i !== index));
  };
  const updateOutcome = (index: number, value: string) => {
    setOutcomeNames(outcomeNames.map((n, i) => (i === index ? value : n)));
  };
  const loadOutcomesFromSelection = async () => {
    setError(null);
    try {
      const { values } = await readSelectedRangeWithValues();
      const parsed = parseOutcomesFromRange(values);
      if (parsed.length === 0) {
        setError("No non-empty outcome names found in the selected range.");
        return;
      }
      setOutcomeNames(parsed);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to read selected range.");
    }
  };

  // Outcomes feed the template build, training, and cell mapping — lock the
  // editor once those have happened so the committed names can't drift.
  const outcomesLocked = sheetCreated || !!initialEvalConfig;

  const handleCreateTemplate = async () => {
    setIsCreatingSheet(true);
    setError(null);
    try {
      const sheetName = `${modelName.substring(0, 20)} Model Def`;
      const config = await createTemplateSheet({
        modelName,
        variables,
        outcomeNames: filledOutcomeNames,
        sheetName,
        formulas,
      });
      setEvalConfig({ ...config, calcMode });
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

  // Bulk-map: take the user's current rectangular selection and assign its
  // cells, in row-major order, to all inputs (or all outcomes) at once. The
  // count must match exactly so cells line up with variables/outcomes in order.
  const handleLoadAllFromSelection = async (type: "input" | "output") => {
    setError(null);
    const expected = type === "input" ? variables.length : filledOutcomeNames.length;
    try {
      const addresses = await readSelectedCellAddresses();
      if (addresses.length !== expected) {
        setError(
          `Selected ${addresses.length} cell(s) but there ${expected === 1 ? "is" : "are"} ${expected} ${type === "input" ? "input" : "outcome"}${expected === 1 ? "" : "s"}. Select exactly one cell per ${type === "input" ? "input, in input order" : "outcome, in outcome order"}.`
        );
        return;
      }
      if (type === "input") {
        setInputCellMap(addresses);
      } else {
        setOutputCellMap(addresses);
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to read selected range.");
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
      outcomeCount: filledOutcomeNames.length,
      inputCells: inputCellMap.map((c) => c.trim()),
      outputCells: outputCellMap.map((c) => c.trim()),
      calcMode,
    };
    setEvalConfig(config);
    setSheetCreated(true);
    setError(null);
  };

  const handleEvaluate = async () => {
    if (!evalConfig || !client) return;

    setIsEvaluating(true);
    setError(null);

    // Apply the current calc-mode selection even if the mapping was
    // confirmed before the user toggled the option.
    const activeConfig: EvalConfig = { ...evalConfig, calcMode };

    try {
      const { outputCases, errors } = await evaluateAllCases(
        activeConfig,
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
        filledOutcomeNames.length,
        outputCases
      );

      // Save cell mappings to _GMOO_State sheet for future re-use
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
          projectId,
          calcMode,
          variables: variables.map((v, i) => ({
            name: v.name,
            type: v.type,
            min: v.min,
            max: v.max,
            inputCell: inputCells[i],
          })),
          outcomes: filledOutcomeNames.map((name, i) => ({
            name,
            outputCell: outputCells[i],
          })),
        });
      } catch {
        // Non-critical — don't fail the evaluation if state save fails
        console.warn("[GMOO] Failed to save state sheet");
      }

      onComplete(trial.id, activeConfig, filledOutcomeNames);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Evaluation failed.");
    } finally {
      setIsEvaluating(false);
    }
  };

  // Require at least one row in each map — `[].every` is vacuously true, so
  // an empty outcomeNames array would otherwise silently flip the badge to
  // "all mapped" and let the user submit a 0-output trial that the API
  // rejects with a cryptic "Output count must be greater than zero".
  const allInputsMapped = inputCellMap.length > 0 && inputCellMap.every((c) => c.trim());
  const allOutputsMapped = outputCellMap.length > 0 && outputCellMap.every((c) => c.trim());

  // Empty inputCases means we landed here from a Resume-from-Project flow
  // where the project DTO didn't ship the (often-large) cases array. There's
  // nothing to evaluate, and submitting an empty trial would silently corrupt
  // the project. Bail out with a hint instead.
  if (inputCases.length === 0) {
    return (
      <div className={styles.container}>
        {headerSlot}
        <Text weight="semibold" size={400}>
          Evaluate Input Cases
        </Text>
        <MessageBar intent="warning">
          <MessageBarBody>
            <MessageBarTitle>No input cases loaded</MessageBarTitle>
            This project doesn&apos;t carry its input cases in the resume payload.
            You can start a new trial by reloading the project&apos;s input cases and
            re-evaluating them (this creates a fresh trial and won&apos;t modify any
            existing trial), use the &quot;Switch trial&quot; selector above to pick a
            trial that&apos;s already been trained, or go back and start a new project.
          </MessageBarBody>
        </MessageBar>
        {reloadError && (
          <MessageBar intent="error">
            <MessageBarBody>{reloadError}</MessageBarBody>
          </MessageBar>
        )}
        <div className={styles.buttonRow}>
          <Button appearance="secondary" onClick={onBack}>
            Back
          </Button>
          {onReloadInputCases && (
            <Button
              appearance="primary"
              onClick={handleReloadInputCases}
              disabled={isReloadingCases}
            >
              {isReloadingCases ? <Spinner size="tiny" /> : "Start a new trial (reload input cases)"}
            </Button>
          )}
        </div>
      </div>
    );
  }

  // Defensive bail when the model has no input variables. Outcomes are now
  // defined on this step, so an empty outcome list is the normal starting state
  // and is handled by the editor below — only missing inputs is a broken state.
  if (variables.length === 0) {
    return (
      <div className={styles.container}>
        {headerSlot}
        <Text weight="semibold" size={400}>
          Evaluate Input Cases
        </Text>
        <MessageBar intent="warning">
          <MessageBarBody>
            <MessageBarTitle>Model definition is incomplete</MessageBarTitle>
            This project has no input variables defined in this workbook. Go back to
            Define Model to set them up before evaluating cases.
          </MessageBarBody>
        </MessageBar>
        <div className={styles.buttonRow}>
          <Button appearance="secondary" onClick={onBack}>
            Back to Define Model
          </Button>
        </div>
      </div>
    );
  }

  return (
    <div className={styles.container}>
      {headerSlot}
      <Text weight="semibold" size={400}>
        Evaluate Input Cases
      </Text>
      <Text size={200}>
        The API generated {inputCases.length} input cases. Your Excel formulas will compute
        the outputs for each case to train the model.
      </Text>

      {/* Outcomes editor — outcomes are defined here (the API only needs them at
          training time). Locked once a template is built or cells are mapped. */}
      <Card>
        <div style={{ padding: "12px", display: "flex", flexDirection: "column", gap: "8px" }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: "8px" }}>
            <Text weight="semibold" size={300}>Outcomes</Text>
            <Button
              icon={<CursorClick20Regular />}
              size="small"
              appearance="subtle"
              onClick={loadOutcomesFromSelection}
              disabled={outcomesLocked}
              title="Populate outcome names from the currently-selected cells (single column; any sheet or workbook)."
            >
              Load from selection
            </Button>
          </div>
          <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
            Name each model output you want to train and optimize against.
          </Text>
          {outcomeNames.map((name, i) => (
            <div key={i} style={{ display: "flex", gap: "8px", alignItems: "center" }}>
              <Input
                size="small"
                value={name}
                onChange={(_, data) => updateOutcome(i, data.value)}
                placeholder={`Outcome ${i + 1}`}
                disabled={outcomesLocked}
                style={{ flexGrow: 1 }}
              />
              <Button
                icon={<Delete20Regular />}
                size="small"
                appearance="subtle"
                onClick={() => removeOutcome(i)}
                disabled={outcomesLocked || outcomeNames.length <= 1}
              />
            </div>
          ))}
          {!outcomesLocked && (
            <Button
              icon={<Add20Regular />}
              size="small"
              appearance="subtle"
              onClick={addOutcome}
              style={{ alignSelf: "flex-start" }}
            >
              Add Outcome
            </Button>
          )}
        </div>
      </Card>

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
            Fill in your outcome formulas in the green Formula cells (column B), referencing
            the blue Current Value cells (column D). The Reference column shows each formula's
            text so you can verify the wiring. Then click "Evaluate Cases & Create New Trial".
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
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: "8px" }}>
              <Text weight="semibold" size={300}>
                Map Input Variables
              </Text>
              <Button
                icon={<CursorClick20Regular />}
                size="small"
                appearance="subtle"
                onClick={() => handleLoadAllFromSelection("input")}
                disabled={isPicking !== null}
                title={`Select a contiguous range of ${variables.length} cell(s) — one per input, in order — then click to map them all at once.`}
              >
                Map all from selection
              </Button>
            </div>
            <Text size={200}>
              Pick each cell individually below, or select a range of {variables.length} cell
              {variables.length === 1 ? "" : "s"} (one per input, in order) and use
              "Map all from selection". Cells can be on different sheets.
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

            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: "8px", marginTop: "8px" }}>
              <Text weight="semibold" size={300}>
                Map Outcomes
              </Text>
              <Button
                icon={<CursorClick20Regular />}
                size="small"
                appearance="subtle"
                onClick={() => handleLoadAllFromSelection("output")}
                disabled={isPicking !== null}
                title={`Select a contiguous range of ${filledOutcomeNames.length} cell(s) — one per outcome, in order — then click to map them all at once.`}
              >
                Map all from selection
              </Button>
            </div>
            <Text size={200}>
              Pick each cell individually below, or select a range of {filledOutcomeNames.length} cell
              {filledOutcomeNames.length === 1 ? "" : "s"} (one per outcome, in order) and use
              "Map all from selection".
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
                {filledOutcomeNames.map((name, i) => (
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
        <div style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
          <div style={{ display: "flex", flexDirection: "column", gap: "2px" }}>
            <Text size={200} weight="semibold">
              Excel calculation per case
            </Text>
            <RadioGroup
              layout="horizontal"
              value={calcMode}
              onChange={(_, data) => setCalcMode(data.value as CalcMode)}
              disabled={isEvaluating}
            >
              <Radio value="recalculate" label="Changed-only (fast)" />
              <Radio value="full" label="Full recalculation" />
            </RadioGroup>
            <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
              Changed-only recalculates just the formulas affected by the input
              cells (like pressing F9) and is recommended. Full recalculates
              every formula in every open workbook each case — much slower on
              large models; use it only if outputs seem stale with changed-only.
            </Text>
          </div>
          <Button
            appearance="primary"
            onClick={handleEvaluate}
            disabled={isEvaluating}
          >
            {isEvaluating ? <Spinner size="tiny" /> : "Evaluate Cases & Create New Trial"}
          </Button>
          <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
            Runs your formulas against all {inputCases.length} input case{inputCases.length === 1 ? "" : "s"} and
            submits the results as a new trial. Existing trials in this project aren't modified.
          </Text>
        </div>
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
