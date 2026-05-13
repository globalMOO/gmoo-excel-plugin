import React, { useCallback, useEffect, useState } from "react";
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
import { ConnectionSetup } from "./components/ConnectionSetup";
import { ModelSetup } from "./components/ModelSetup";
import { CaseEvaluation } from "./components/CaseEvaluation";
import { ObjectiveSetup } from "./components/ObjectiveSetup";
import { OptimizationRunner } from "./components/OptimizationRunner";
import { ResultsSummary } from "./components/ResultsSummary";
import { MultiSolvePanel } from "./components/MultiSolvePanel";
import { useConnections } from "./hooks/useConnections";
import { useGmooClient } from "./hooks/useGmooClient";
import { useWorkbookState } from "./hooks/useWorkbookState";
import { useOptimization } from "./hooks/useOptimization";
import { WizardStep } from "./types/workbookState";
import type { EvalConfig, VsmeStateData } from "./services/excelService";
import { loadStateSheet } from "./services/excelService";
import type { InputVariable } from "./types/workbookState";
import type { ObjectiveRowData } from "./components/ObjectiveSetup";
import type { Example } from "./examples";
import {
  parseActivationFromUrl,
  exchangeActivation,
  applyActivation,
  clearActivationFromUrl,
  ActivationError,
} from "./services/activationService";

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

type Banner = { intent: "success" | "error" | "info"; title: string; body?: string } | null;

const App: React.FC = () => {
  const styles = useStyles();
  const { state, isLoaded: isStateLoaded, updateState, resetState } = useWorkbookState();

  const setActiveConnectionIdInState = useCallback(
    async (id: string | null) => {
      await updateState({ activeConnectionId: id });
    },
    [updateState]
  );

  const {
    connections,
    activeConnection,
    isLoading: isLoadingConnections,
    setActive,
    add,
    update,
    remove,
    refresh: refreshConnections,
  } = useConnections({
    activeConnectionId: state.activeConnectionId,
    setActiveConnectionId: setActiveConnectionIdInState,
  });

  const client = useGmooClient(activeConnection);

  // Activation flow: parse URL on first launch and run the exchange. While the
  // exchange is in flight we render a full-pane spinner. Errors fall through
  // to the wizard with a banner. Success applies the connection, sets it
  // active, and proceeds.
  const [activationStatus, setActivationStatus] = useState<
    | { kind: "idle" }
    | { kind: "exchanging"; hostname: string }
    | { kind: "done" }
  >(() => {
    const params = parseActivationFromUrl();
    if (!params) return { kind: "done" };
    let host = params.srv;
    try {
      host = new URL(params.srv).hostname;
    } catch {
      // keep raw srv
    }
    return { kind: "exchanging", hostname: host };
  });
  const [banner, setBanner] = useState<Banner>(null);

  useEffect(() => {
    if (activationStatus.kind !== "exchanging") return;
    const params = parseActivationFromUrl();
    if (!params) {
      setActivationStatus({ kind: "done" });
      return;
    }
    let cancelled = false;
    (async () => {
      try {
        const result = await exchangeActivation(params.srv, params.token);
        const conn = await applyActivation(result, params.label);
        await refreshConnections();
        await setActiveConnectionIdInState(conn.id);
        if (!cancelled) {
          setBanner({
            intent: "success",
            title: "Connected",
            body: `Activated "${conn.label}".`,
          });
        }
      } catch (err) {
        if (!cancelled) {
          const msg =
            err instanceof ActivationError
              ? err.message
              : err instanceof Error
              ? err.message
              : "Activation failed.";
          setBanner({ intent: "error", title: "Activation failed", body: msg });
        }
      } finally {
        clearActivationFromUrl();
        if (!cancelled) setActivationStatus({ kind: "done" });
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [activationStatus.kind, refreshConnections, setActiveConnectionIdInState]);

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

  if (isLoadingConnections || !isStateLoaded) {
    return (
      <FluentProvider theme={webLightTheme}>
        <div className={styles.loading}>
          <Spinner label="Loading..." />
        </div>
      </FluentProvider>
    );
  }

  if (activationStatus.kind === "exchanging") {
    return (
      <FluentProvider theme={webLightTheme}>
        <div className={styles.loading}>
          <Spinner label={`Connecting to ${activationStatus.hostname}…`} />
        </div>
      </FluentProvider>
    );
  }

  const currentStep = state.wizardStep as WizardStep;

  const renderStep = () => {
    switch (currentStep) {
      case WizardStep.Connect:
        return (
          <ConnectionSetup
            connections={connections}
            activeConnection={activeConnection}
            onSetActive={setActive}
            onAdd={add}
            onUpdate={update}
            onDelete={remove}
            onNext={() => {
              if (activeConnection) {
                updateState({
                  apiKeyHint: activeConnection.apiKey
                    ? `...${activeConnection.apiKey.slice(-4)}`
                    : "",
                });
              }
              goToStep(WizardStep.DefineModel);
            }}
            banner={banner}
            onDismissBanner={() => setBanner(null)}
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
            onMultiSolve={() => goToStep(WizardStep.MultiSolve)}
          />
        );

      case WizardStep.MultiSolve:
        return (
          <MultiSolvePanel
            client={client}
            trialId={state.trialId}
            evalConfig={evalConfig}
            inputVariables={state.inputVariables}
            outcomeNames={state.outcomeNames}
            objectiveRows={savedObjectives}
            onBack={() => goToStep(WizardStep.Results)}
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
