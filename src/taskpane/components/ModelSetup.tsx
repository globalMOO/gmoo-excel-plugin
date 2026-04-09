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
  Card,
  Caption1,
  ProgressBar,
} from "@fluentui/react-components";
import { Add20Regular, Delete20Regular, BookTemplate20Regular, ChevronDown20Regular, ChevronUp20Regular } from "@fluentui/react-icons";
import type { GmooClient } from "../services/gmooApi";
import type { InputVariable } from "../types/workbookState";
import type { EvalConfig } from "../services/excelService";
import { createTemplateSheet, createExampleSheets } from "../services/excelService";
import { InputType } from "../types/gmoo";
import { EXAMPLES, isMultiSheetExample, type Example } from "../examples";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    padding: "16px",
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  row: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
  },
  smallInput: {
    width: "80px",
  },
  nameInput: {
    width: "100px",
  },
  typeDropdown: {
    minWidth: "90px",
  },
  buttonRow: {
    display: "flex",
    justifyContent: "space-between",
    marginTop: "12px",
  },
  outcomeRow: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
  },
  exampleToggle: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    cursor: "pointer",
    padding: "8px 12px",
    borderRadius: "6px",
    backgroundColor: tokens.colorBrandBackground2,
    border: `1px solid ${tokens.colorBrandStroke1}`,
    "&:hover": {
      backgroundColor: tokens.colorBrandBackground2Hover,
    },
  },
  exampleCard: {
    cursor: "pointer",
    padding: "8px 12px",
    "&:hover": {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  exampleCardSelected: {
    cursor: "pointer",
    padding: "8px 12px",
    backgroundColor: tokens.colorBrandBackground2,
    border: `1px solid ${tokens.colorBrandStroke1}`,
  },
  exampleList: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    marginTop: "6px",
  },
  setupOverlay: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "32px 16px",
    alignItems: "center",
    textAlign: "center",
  },
  progressSection: {
    width: "100%",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  complexityBadge: {
    fontSize: "10px",
    padding: "1px 6px",
    borderRadius: "4px",
    marginLeft: "8px",
    fontWeight: "normal",
  },
});

export interface ModelSetupCompleteData {
  modelId: number;
  projectId: number;
  modelName: string;
  inputVariables: InputVariable[];
  outcomeNames: string[];
  inputCases: number[][];
  selectedExample?: Example;
  /** Present when an example ran the full auto-setup pipeline */
  trialId?: number;
  evalConfig?: EvalConfig;
}

interface ModelSetupProps {
  client: GmooClient | null;
  onComplete: (data: ModelSetupCompleteData) => void;
  onBack: () => void;
  initialVariables?: InputVariable[];
  initialOutcomes?: string[];
  initialModelName?: string;
}

const INPUT_TYPE_OPTIONS = [
  { value: InputType.Float, label: "Float" },
  { value: InputType.Integer, label: "Integer" },
  { value: InputType.Boolean, label: "Boolean" },
  { value: InputType.Category, label: "Category" },
];

// Local editing state uses strings for min/max so typing "-1" works naturally
interface VariableRow {
  name: string;
  type: string;
  min: string;
  max: string;
  categories?: string[];
}

function toVariableRows(vars: InputVariable[]): VariableRow[] {
  return vars.map((v) => ({ ...v, min: String(v.min), max: String(v.max) }));
}

function toInputVariables(rows: VariableRow[]): InputVariable[] {
  return rows.map((r) => ({
    ...r,
    min: parseFloat(r.min) || 0,
    max: parseFloat(r.max) || 0,
  }));
}

const COMPLEXITY_COLORS: Record<string, string> = {
  beginner: "#107C10",
  intermediate: "#CA5010",
  advanced: "#D13438",
};

export const ModelSetup: React.FC<ModelSetupProps> = ({
  client,
  onComplete,
  onBack,
  initialVariables,
  initialOutcomes,
  initialModelName,
}) => {
  const styles = useStyles();

  const [modelName, setModelName] = useState(
    initialModelName || `test-${new Date().toISOString().slice(5, 16).replace(/[T:]/g, "-")}`
  );
  const [description, setDescription] = useState("");
  const [variables, setVariables] = useState<VariableRow[]>(
    initialVariables ? toVariableRows(initialVariables) : [
      { name: "", type: InputType.Float, min: "0", max: "1" },
      { name: "", type: InputType.Float, min: "0", max: "1" },
    ]
  );
  const [outcomeNames, setOutcomeNames] = useState<string[]>(
    initialOutcomes ?? [""]
  );
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [showExamples, setShowExamples] = useState(false);
  const [selectedExample, setSelectedExample] = useState<Example | null>(null);

  // Auto-setup progress state
  const [isAutoSetup, setIsAutoSetup] = useState(false);
  const [setupStep, setSetupStep] = useState("");
  const [setupProgress, setSetupProgress] = useState<{ current: number; total: number } | null>(null);

  /**
   * Auto-setup: create model + project via API, build spreadsheet(s) in Excel,
   * then hand off to EvaluateCases so the user can see the sheet and run training.
   */
  const loadExample = async (example: Example) => {
    if (!client) {
      setError("Not connected to API.");
      return;
    }

    setSelectedExample(example);
    setShowExamples(false);
    setIsAutoSetup(true);
    setError(null);

    try {
      // Step 1: Create model + project via API
      setSetupStep("Creating model...");
      const name = `${example.name.substring(0, 20)}-${new Date().toISOString().slice(5, 16).replace(/[T:]/g, "-")}`;
      const desc = `Example: ${example.name}`;
      const paddedDesc = desc.length < 8 ? desc.padEnd(8, " ") : desc;
      const model = await client.createModel(name, paddedDesc);

      const project = await client.createProject(
        model.id,
        name,
        example.variables.length,
        example.variables.map((v) => v.min),
        example.variables.map((v) => v.max),
        example.variables.map((v) => v.type),
        example.variables.flatMap((v) => v.categories ?? [])
      );

      // Step 2: Build spreadsheet(s)
      setSetupStep("Building spreadsheet...");
      let evalConfig: EvalConfig;

      if (isMultiSheetExample(example)) {
        // Multi-sheet mode
        evalConfig = await createExampleSheets(
          example.setup.sheets!,
          example.setup.inputCells!,
          example.setup.outputCells!,
          example.variables.length,
          example.outcomeNames.length
        );
      } else {
        // Simple template mode
        const sheetName = `${name.substring(0, 20)} Model Def`;
        evalConfig = await createTemplateSheet({
          modelName: name,
          variables: example.variables,
          outcomeNames: example.outcomeNames,
          sheetName,
          formulas: example.setup.formulas,
        });
      }

      // Done — hand off to EvaluateCases so user can see the sheet and run training
      onComplete({
        modelId: model.id,
        projectId: project.id,
        modelName: name,
        inputVariables: example.variables,
        outcomeNames: example.outcomeNames,
        inputCases: project.inputCases,
        selectedExample: example,
        evalConfig,
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : "Example setup failed.");
      setIsAutoSetup(false);
    }
  };

  const addVariable = () => {
    setVariables([...variables, { name: "", type: InputType.Float, min: "0", max: "1" }]);
  };

  const removeVariable = (index: number) => {
    if (variables.length <= 2) return;
    setVariables(variables.filter((_, i) => i !== index));
  };

  const updateVariable = (index: number, field: keyof VariableRow, value: string) => {
    setVariables(
      variables.map((v, i) =>
        i === index ? { ...v, [field]: value } : v
      )
    );
  };

  const addOutcome = () => {
    setOutcomeNames([...outcomeNames, ""]);
  };

  const removeOutcome = (index: number) => {
    if (outcomeNames.length <= 1) return;
    setOutcomeNames(outcomeNames.filter((_, i) => i !== index));
  };

  const validate = (): string | null => {
    if (!modelName.trim()) return "Model name is required.";
    if (modelName.trim().length < 4) return "Model name must be at least 4 characters.";
    if (variables.length < 2) return "At least 2 input variables are required.";
    for (let i = 0; i < variables.length; i++) {
      if (!variables[i].name.trim()) return `Variable ${i + 1} needs a name.`;
      const t = variables[i].type;
      const minVal = parseFloat(variables[i].min);
      const maxVal = parseFloat(variables[i].max);
      if (t !== "boolean" && t !== "category" && (isNaN(minVal) || isNaN(maxVal) || minVal >= maxVal)) {
        return `Variable "${variables[i].name}": min must be less than max.`;
      }
    }
    if (outcomeNames.length === 0) return "At least one outcome is required.";
    for (let i = 0; i < outcomeNames.length; i++) {
      if (!outcomeNames[i].trim()) return `Outcome ${i + 1} needs a name.`;
    }
    return null;
  };

  const handleSubmit = async () => {
    const validationError = validate();
    if (validationError) {
      setError(validationError);
      return;
    }
    if (!client) {
      setError("Not connected to API.");
      return;
    }

    setIsSubmitting(true);
    setError(null);

    try {
      const desc = description.trim() || `VSME model: ${modelName.trim()}`;
      const paddedDesc = desc.length < 8 ? desc.padEnd(8, " ") : desc;
      const model = await client.createModel(modelName.trim(), paddedDesc);
      const inputVars = toInputVariables(variables);
      const project = await client.createProject(
        model.id,
        modelName.trim(),
        inputVars.length,
        inputVars.map((v) => v.min),
        inputVars.map((v) => v.max),
        inputVars.map((v) => v.type),
        inputVars.flatMap((v) => v.categories ?? [])
      );

      onComplete({
        modelId: model.id,
        projectId: project.id,
        modelName: modelName.trim(),
        inputVariables: inputVars,
        outcomeNames: outcomeNames.map((n) => n.trim()),
        inputCases: project.inputCases,
        selectedExample: selectedExample ?? undefined,
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to create model.");
    } finally {
      setIsSubmitting(false);
    }
  };

  // --- Auto-setup progress overlay ---
  if (isAutoSetup) {
    return (
      <div className={styles.container}>
        <div className={styles.setupOverlay}>
          <Spinner size="large" />
          <Text weight="semibold" size={400}>
            Setting up: {selectedExample?.name}
          </Text>
          <div className={styles.progressSection}>
            <Text size={200}>{setupStep}</Text>
            {setupProgress && (
              <>
                <ProgressBar
                  value={setupProgress.current / setupProgress.total}
                  max={1}
                />
                <Text size={100}>
                  Case {setupProgress.current} of {setupProgress.total}
                </Text>
              </>
            )}
          </div>
        </div>
        {error && (
          <MessageBar intent="error">
            <MessageBarBody>{error}</MessageBarBody>
          </MessageBar>
        )}
      </div>
    );
  }

  // --- Normal form ---
  return (
    <div className={styles.container}>
      <Text weight="semibold" size={400}>
        Define Model
      </Text>

      <div>
        <div
          className={styles.exampleToggle}
          onClick={() => setShowExamples(!showExamples)}
          role="button"
          tabIndex={0}
        >
          <BookTemplate20Regular />
          <Text size={200} weight="semibold">
            Load an example
          </Text>
          {showExamples ? <ChevronUp20Regular /> : <ChevronDown20Regular />}
        </div>
        {showExamples && (
          <div className={styles.exampleList}>
            {EXAMPLES.map((ex) => (
              <Card
                key={ex.id}
                className={styles.exampleCard}
                size="small"
                onClick={() => loadExample(ex)}
              >
                <div style={{ display: "flex", alignItems: "center" }}>
                  <Text size={200} weight="semibold">{ex.name}</Text>
                  {ex.complexity && (
                    <span
                      className={styles.complexityBadge}
                      style={{
                        color: COMPLEXITY_COLORS[ex.complexity] ?? tokens.colorNeutralForeground2,
                        border: `1px solid ${COMPLEXITY_COLORS[ex.complexity] ?? tokens.colorNeutralStroke1}`,
                      }}
                    >
                      {ex.complexity}
                    </span>
                  )}
                </div>
                <Caption1>{ex.description}</Caption1>
              </Card>
            ))}
          </div>
        )}
      </div>

      <div className={styles.section}>
        <Text size={200}>Model Name</Text>
        <Input
          value={modelName}
          onChange={(_, data) => setModelName(data.value)}
          placeholder="My VSME Model (min 4 chars)"
        />
        <Text size={200}>Description (optional)</Text>
        <Input
          value={description}
          onChange={(_, data) => setDescription(data.value)}
          placeholder="Optional description"
        />
      </div>

      <div className={styles.section}>
        <Text weight="semibold" size={300}>
          Input Variables
        </Text>
        <Table size="extra-small">
          <TableHeader>
            <TableRow>
              <TableHeaderCell>Name</TableHeaderCell>
              <TableHeaderCell>Type</TableHeaderCell>
              <TableHeaderCell>Min</TableHeaderCell>
              <TableHeaderCell>Max</TableHeaderCell>
              <TableHeaderCell></TableHeaderCell>
            </TableRow>
          </TableHeader>
          <TableBody>
            {variables.map((v, i) => (
              <TableRow key={i}>
                <TableCell>
                  <Input
                    className={styles.nameInput}
                    size="small"
                    value={v.name}
                    onChange={(_, data) => updateVariable(i, "name", data.value)}
                    placeholder={`Var ${i + 1}`}
                  />
                </TableCell>
                <TableCell>
                  <Dropdown
                    className={styles.typeDropdown}
                    size="small"
                    value={v.type}
                    selectedOptions={[v.type]}
                    onOptionSelect={(_, data) =>
                      updateVariable(i, "type", data.optionValue ?? InputType.Float)
                    }
                  >
                    {INPUT_TYPE_OPTIONS.map((opt) => (
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
                    value={v.min}
                    onChange={(_, data) => updateVariable(i, "min", data.value)}
                  />
                </TableCell>
                <TableCell>
                  <Input
                    className={styles.smallInput}
                    size="small"
                    value={v.max}
                    onChange={(_, data) => updateVariable(i, "max", data.value)}
                  />
                </TableCell>
                <TableCell>
                  <Button
                    icon={<Delete20Regular />}
                    size="small"
                    appearance="subtle"
                    onClick={() => removeVariable(i)}
                    disabled={variables.length <= 2}
                  />
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
        <Button
          icon={<Add20Regular />}
          size="small"
          appearance="subtle"
          onClick={addVariable}
          disabled={variables.length >= 200}
        >
          Add Variable
        </Button>
      </div>

      <div className={styles.section}>
        <Text weight="semibold" size={300}>
          Outcomes
        </Text>
        {outcomeNames.map((name, i) => (
          <div key={i} className={styles.outcomeRow}>
            <Input
              size="small"
              value={name}
              onChange={(_, data) => {
                const updated = [...outcomeNames];
                updated[i] = data.value;
                setOutcomeNames(updated);
              }}
              placeholder={`Outcome ${i + 1}`}
              style={{ flexGrow: 1 }}
            />
            <Button
              icon={<Delete20Regular />}
              size="small"
              appearance="subtle"
              onClick={() => removeOutcome(i)}
              disabled={outcomeNames.length <= 1}
            />
          </div>
        ))}
        <Button
          icon={<Add20Regular />}
          size="small"
          appearance="subtle"
          onClick={addOutcome}
        >
          Add Outcome
        </Button>
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
          {isSubmitting ? <Spinner size="tiny" /> : "Create Model & Continue"}
        </Button>
      </div>
    </div>
  );
};
