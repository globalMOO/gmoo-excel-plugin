export interface InputVariable {
  name: string;
  type: string;
  min: number;
  max: number;
  categories?: string[];
}

export interface WorkbookState {
  apiKeyHint: string;
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
  inputRangeAddress: string | null;
  outputRangeAddress: string | null;
  wizardStep: number;
}

export const DEFAULT_WORKBOOK_STATE: WorkbookState = {
  apiKeyHint: "",
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
  inputRangeAddress: null,
  outputRangeAddress: null,
  wizardStep: 0,
};

export enum WizardStep {
  Connect = 0,
  DefineModel = 1,
  EvaluateCases = 2,
  SetObjectives = 3,
  Optimize = 4,
  Results = 5,
}

export const WIZARD_STEP_LABELS: Record<WizardStep, string> = {
  [WizardStep.Connect]: "Connect",
  [WizardStep.DefineModel]: "Define Model",
  [WizardStep.EvaluateCases]: "Evaluate Cases",
  [WizardStep.SetObjectives]: "Set Objectives",
  [WizardStep.Optimize]: "Optimize",
  [WizardStep.Results]: "Results",
};
