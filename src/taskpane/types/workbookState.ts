export interface InputVariable {
  name: string;
  type: string;
  min: number;
  max: number;
  categories?: string[];
}

/** Bumped when the WizardStep enum is renumbered. Used by stateStore to migrate
 *  persisted values; current schema is 2 (after the Resume step was inserted). */
export const STATE_SCHEMA_VERSION = 2;

export interface WorkbookState {
  schemaVersion?: number;
  apiKeyHint: string;
  /** ID of the Connection (from connectionsService) selected for this workbook. */
  activeConnectionId: string | null;
  modelId: number | null;
  modelName: string;
  projectId: number | null;
  trialId: number | null;
  objectiveId: number | null;
  inputVariables: InputVariable[];
  outcomeNames: string[];
  inputCases: number[][] | null;
  formulaMode: "template" | "existing";
  sheetName: string | null;
  wizardStep: number;
}

export const DEFAULT_WORKBOOK_STATE: WorkbookState = {
  schemaVersion: STATE_SCHEMA_VERSION,
  apiKeyHint: "",
  activeConnectionId: null,
  modelId: null,
  modelName: "",
  projectId: null,
  trialId: null,
  objectiveId: null,
  inputVariables: [],
  outcomeNames: [],
  inputCases: null,
  formulaMode: "template",
  sheetName: null,
  wizardStep: 0,
};

export enum WizardStep {
  Connect = 0,
  Resume = 1,
  DefineModel = 2,
  EvaluateCases = 3,
  SetObjectives = 4,
  Optimize = 5,
  Results = 6,
  MultiSolve = 7,
}

export const WIZARD_STEP_LABELS: Record<WizardStep, string> = {
  [WizardStep.Connect]: "Connect",
  [WizardStep.Resume]: "Resume / New",
  [WizardStep.DefineModel]: "Define Model",
  [WizardStep.EvaluateCases]: "Evaluate Cases",
  [WizardStep.SetObjectives]: "Set Objectives",
  [WizardStep.Optimize]: "Optimize",
  [WizardStep.Results]: "Results",
  [WizardStep.MultiSolve]: "Multi-Solve",
};
