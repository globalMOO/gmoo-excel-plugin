import React, { useCallback, useEffect } from "react";
import { BUILD_TIME } from "./buildInfo";
import {
  FluentProvider,
  webLightTheme,
  makeStyles,
  tokens,
  Text,
  Spinner,
  Button,
} from "@fluentui/react-components";
import { WizardStepper } from "./components/WizardStepper";
import { ApiKeySetup } from "./components/ApiKeySetup";
import { ModelSetup } from "./components/ModelSetup";
import { CaseEvaluation } from "./components/CaseEvaluation";
import { ObjectiveSetup } from "./components/ObjectiveSetup";
import { OptimizationRunner } from "./components/OptimizationRunner";
import { ResultsSummary } from "./components/ResultsSummary";
import { useApiKey } from "./hooks/useApiKey";
import { useGmooClient } from "./hooks/useGmooClient";
import { useWorkbookState } from "./hooks/useWorkbookState";
import { useOptimization } from "./hooks/useOptimization";
import { WizardStep } from "./types/workbookState";
import type { EvalConfig, VsmeStateData } from "./services/excelService";
import { loadStateSheet } from "./services/excelService";
import type { InputVariable } from "./types/workbookState";
import type { ObjectiveRowData } from "./components/ObjectiveSetup";
import type { Example } from "./examples";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  header: {
    padding: "12px 16px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorBrandBackground,
  },
  headerTitle: {
    color: tokens.colorNeutralForegroundOnBrand,
  },
  content: {
    flexGrow: 1,
    overflowY: "auto",
  },
  loading: {
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    height: "100vh",
  },
  errorContainer: {
    padding: "16px",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    alignItems: "center",
  },
});

const App: React.FC = () => {
  const styles = useStyles();
  const { apiKey, apiUrl, setApiKey, setApiUrl, isLoading: isLoadingKey, DEFAULT_API_URL } = useApiKey();
  const client = useGmooClient(apiKey, apiUrl);
  const { state, isLoaded: isStateLoaded, updateState, resetState } = useWorkbookState();

  // Keep evalConfig in React state (not persisted — re-created on template step)
  const [evalConfig, setEvalConfig] = React.useState<EvalConfig | null>(null);

  // Keep objective row values in React state so they survive back-navigation
  const [savedObjectives, setSavedObjectives] = React.useState<ObjectiveRowData[] | null>(null);

  // Selected example (carries formulas for template + default objectives)
  const [selectedExample, setSelectedExample] = React.useState<Example | null>(null);

  // Pre-fill from _VSME_State sheet if it exists
  const [savedStateData, setSavedStateData] = React.useState<VsmeStateData | null>(null);

  useEffect(() => {
    loadStateSheet().then((data) => {
      if (data) setSavedStateData(data);
    }).catch(() => {});
  }, []);

  const optimization = useOptimization(client, state.objectiveId, evalConfig);

  const goToStep = useCallback(
    (step: WizardStep) => {
      updateState({ wizardStep: step });
    },
    [updateState]
  );

  if (isLoadingKey || !isStateLoaded) {
    return (
      <FluentProvider theme={webLightTheme}>
        <div className={styles.loading}>
          <Spinner label="Loading..." />
        </div>
      </FluentProvider>
    );
  }

  const currentStep = state.wizardStep as WizardStep;

  const renderStep = () => {
    switch (currentStep) {
      case WizardStep.Connect:
        return (
          <ApiKeySetup
            apiKey={apiKey}
            apiUrl={apiUrl}
            defaultApiUrl={DEFAULT_API_URL}
            onApiKeyChange={async (key) => {
              await setApiKey(key);
              updateState({ apiKeyHint: key ? `...${key.slice(-4)}` : "" });
            }}
            onApiUrlChange={setApiUrl}
            onNext={() => goToStep(WizardStep.DefineModel)}
          />
        );

      case WizardStep.DefineModel:
        return (
          <ModelSetup
            client={client}
            initialModelName={state.modelName}
            initialVariables={
              state.inputVariables.length > 0
                ? state.inputVariables
                : savedStateData?.variables.map((v) => ({
                    name: v.name, type: v.type, min: v.min, max: v.max,
                  }))
            }
            initialOutcomes={
              state.outcomeNames.length > 0
                ? state.outcomeNames
                : savedStateData?.outcomes.map((o) => o.name)
            }
            onComplete={(data) => {
              setSavedObjectives(null);
              setSelectedExample(data.selectedExample ?? null);

              if (data.evalConfig) {
                // Example auto-setup built the spreadsheet — skip to EvaluateCases with sheet ready
                setEvalConfig(data.evalConfig);
              }

              updateState({
                modelId: data.modelId,
                projectId: data.projectId,
                modelName: data.modelName,
                inputVariables: data.inputVariables,
                outcomeNames: data.outcomeNames,
                inputCases: data.inputCases,
                wizardStep: WizardStep.EvaluateCases,
              });
            }}
            onBack={() => goToStep(WizardStep.Connect)}
          />
        );

      case WizardStep.EvaluateCases:
        return (
          <CaseEvaluation
            client={client}
            modelName={state.modelName}
            projectId={state.projectId!}
            variables={state.inputVariables}
            outcomeNames={state.outcomeNames}
            inputCases={state.inputCases ?? []}
            formulas={selectedExample?.setup.formulas}
            initialEvalConfig={evalConfig ?? undefined}
            onComplete={(trialId, config) => {
              setEvalConfig(config);
              updateState({
                trialId,
                sheetName: config.sheetName,
                wizardStep: WizardStep.SetObjectives,
              });
            }}
            onBack={() => goToStep(WizardStep.DefineModel)}
          />
        );

      case WizardStep.SetObjectives:
        return (
          <ObjectiveSetup
            client={client}
            trialId={state.trialId!}
            outcomeNames={state.outcomeNames}
            inputCases={state.inputCases ?? []}
            evalConfig={evalConfig}
            initialObjectives={savedObjectives ?? undefined}
            exampleObjectives={selectedExample?.objectives}
            onComplete={(objectiveId, objectiveRows) => {
              setSavedObjectives(objectiveRows);
              updateState({
                objectiveId,
                wizardStep: WizardStep.Optimize,
              });
            }}
            onBack={() => goToStep(WizardStep.EvaluateCases)}
          />
        );

      case WizardStep.Optimize:
        return (
          <OptimizationRunner
            state={optimization}
            onRun={(max) => optimization.run(max)}
            onStop={optimization.stop}
            onRunSingle={() => optimization.runSingleIteration()}
            onNext={() => goToStep(WizardStep.Results)}
            onBack={() => {
              optimization.reset();
              goToStep(WizardStep.SetObjectives);
            }}
          />
        );

      case WizardStep.Results:
        return (
          <ResultsSummary
            iterations={optimization.iterations}
            inputVariableNames={state.inputVariables.map((v: InputVariable) => v.name)}
            outcomeNames={state.outcomeNames}
            onStartOver={async () => {
              optimization.reset();
              setSavedObjectives(null);
              setSelectedExample(null);
              await resetState();
              goToStep(WizardStep.Connect);
            }}
          />
        );

      default:
        return (
          <div className={styles.errorContainer}>
            <Text>Unknown step. Please start over.</Text>
            <Button onClick={() => goToStep(WizardStep.Connect)}>Start Over</Button>
          </div>
        );
    }
  };

  return (
    <FluentProvider theme={webLightTheme}>
      <div className={styles.root}>
        <div className={styles.header}>
          <Text className={styles.headerTitle} weight="semibold" size={400}>
            VSME - globalMOO
          </Text>
          <Text className={styles.headerTitle} size={100}>
            Build: {BUILD_TIME}
          </Text>
        </div>
        <WizardStepper currentStep={currentStep} />
        <div className={styles.content}>{renderStep()}</div>
      </div>
    </FluentProvider>
  );
};

export default App;
