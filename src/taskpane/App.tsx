import React, { useCallback, useEffect, useRef, useState } from "react";
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
import { ResumePicker, type ResumeSelection } from "./components/ResumePicker";
import { PickExistingBar } from "./components/PickExistingBar";
import { useConnections } from "./hooks/useConnections";
import { useGmooClient } from "./hooks/useGmooClient";
import { useWorkbookState } from "./hooks/useWorkbookState";
import { useOptimization } from "./hooks/useOptimization";
import { useAliases } from "./hooks/useAliases";
import { useProjectCatalog } from "./hooks/useProjectCatalog";
import { WizardStep, type WorkbookState } from "./types/workbookState";
import type { Project, Trial } from "./types/gmoo";
import { ObjectiveType } from "./types/gmoo";
import type { EvalConfig, GmooStateData } from "./services/excelService";
import { loadStateSheet } from "./services/excelService";
import { resolveResumeNames } from "./services/resumeNames";
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

  // Alias registry + project catalog for the Resume picker and the inline
  // PickExistingBar shown on subsequent wizard steps.
  const aliases = useAliases(activeConnection?.id ?? null);
  const catalog = useProjectCatalog(client);

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

  // When the active connection truly switches from one identifiable connection
  // to another (X → Y), the workbook entity ids — modelId, projectId, trialId,
  // objectiveId — refer to the *previous* tenant and would be meaningless or
  // hit 404s. Clear them and bounce the user back to Resume so they can pick
  // afresh.
  //
  // The ref tracks three transitions distinctly:
  //   undefined → anything : first-time capture, no wipe (loading a workbook).
  //   null      → X        : "adoption", e.g. user just added their first
  //                          connection while sitting on Connect. State is
  //                          already empty; teleporting them to Resume would
  //                          yank them out of the connection editor.
  //   X         → Y | null : actual switch (or last connection deleted) —
  //                          stale ids; wipe and send to Resume.
  const prevConnectionIdRef = useRef<string | null | undefined>(undefined);
  useEffect(() => {
    if (isLoadingConnections || !isStateLoaded) return;
    const currentId = activeConnection?.id ?? null;
    const prevId = prevConnectionIdRef.current;
    if (prevId === undefined || prevId === currentId) {
      prevConnectionIdRef.current = currentId;
      return;
    }
    if (prevId === null) {
      // Adoption — first connection added. Nothing to reset.
      prevConnectionIdRef.current = currentId;
      return;
    }
    prevConnectionIdRef.current = currentId;
    setEvalConfig(null);
    setSavedObjectives(null);
    setSelectedExample(null);
    void updateState({
      modelId: null,
      projectId: null,
      trialId: null,
      objectiveId: null,
      inputCases: null,
      inputVariables: [],
      outcomeNames: [],
      modelName: "",
      sheetName: null,
      wizardStep: WizardStep.Resume,
    });
  }, [activeConnection?.id, isLoadingConnections, isStateLoaded, updateState]);

  // Pre-fill from _GMOO_State sheet if it exists.
  // The sheet may carry a tagged projectId (newer saves) — when present we
  // require it to match the active project before rehydrating, so a workbook
  // that has hosted multiple projects over its lifetime doesn't silently feed
  // stale cell mappings to a different project. For pre-tag sheets we fall
  // back to "rehydrate for whichever project was active at app mount".
  const [savedStateData, setSavedStateData] = React.useState<GmooStateData | null>(null);
  const initialProjectIdRef = useRef<number | null | undefined>(undefined);

  useEffect(() => {
    loadStateSheet().then((data) => {
      if (data) setSavedStateData(data);
    }).catch(() => {});
  }, []);

  useEffect(() => {
    if (!isStateLoaded) return;
    if (initialProjectIdRef.current === undefined) {
      initialProjectIdRef.current = state.projectId;
    }
  }, [isStateLoaded, state.projectId]);

  // When the user reopens a workbook, evalConfig (React-only state) is null
  // even though the _GMOO_State sheet on disk has the cell mappings. Without
  // rehydration the user lands on their last saved step (often Optimize) and
  // hits "No spreadsheet bound" with no path forward but to back-navigate
  // through three steps and re-confirm. Rebuild here when the saved-sheet
  // shape still matches the live state, so reopen-and-resume just works.
  //
  // Matching strategy: prefer name match (resilient to reordering), fall back
  // to positional match when names diverge (resilient to in-place renames).
  // Bail only when both fail — the saved sheet truly doesn't describe the
  // current variables and we'd be feeding bad cell addresses to the evaluator.
  useEffect(() => {
    if (!isStateLoaded || evalConfig || !savedStateData) return;
    if (state.inputVariables.length === 0 || state.outcomeNames.length === 0) return;
    // Project-tag guard:
    //   - Tagged sheet: only rehydrate when the tag matches state.projectId.
    //   - Untagged sheet (legacy): only rehydrate for the project that was
    //     active at app mount, since we have no other way to know who owns it.
    if (typeof savedStateData.projectId === "number") {
      if (savedStateData.projectId !== state.projectId) return;
    } else if (
      initialProjectIdRef.current === undefined ||
      state.projectId !== initialProjectIdRef.current
    ) {
      return;
    }

    const sameInputCount =
      savedStateData.variables.length === state.inputVariables.length;
    const sameOutputCount =
      savedStateData.outcomes.length === state.outcomeNames.length;

    const inputCells: string[] = [];
    for (let i = 0; i < state.inputVariables.length; i++) {
      const v = state.inputVariables[i];
      const byName = savedStateData.variables.find((sv) => sv.name === v.name);
      const byIndex = sameInputCount ? savedStateData.variables[i] : undefined;
      const cell = byName?.inputCell ?? byIndex?.inputCell;
      if (!cell) return;
      inputCells.push(cell);
    }
    const outputCells: string[] = [];
    for (let i = 0; i < state.outcomeNames.length; i++) {
      const name = state.outcomeNames[i];
      const byName = savedStateData.outcomes.find((so) => so.name === name);
      const byIndex = sameOutputCount ? savedStateData.outcomes[i] : undefined;
      const cell = byName?.outputCell ?? byIndex?.outputCell;
      if (!cell) return;
      outputCells.push(cell);
    }
    setEvalConfig({
      variableCount: state.inputVariables.length,
      outcomeCount: state.outcomeNames.length,
      inputCells,
      outputCells,
    });
  }, [isStateLoaded, evalConfig, savedStateData, state.inputVariables, state.outcomeNames]);

  // savedObjectives is React-only and feeds both the SetObjectives form and
  // MultiSolvePanel. Reopening a workbook or switching objectives via
  // PickExistingBar leaves it null/stale. Fetch the objective once per id
  // and reconstruct the form rows from the API DTO so Back-navigation and
  // MultiSolve both show the right values.
  const loadedObjectiveForRowsRef = useRef<number | null>(null);
  useEffect(() => {
    if (!isStateLoaded || !client) return;
    const id = state.objectiveId;
    if (!id) {
      loadedObjectiveForRowsRef.current = null;
      return;
    }
    if (loadedObjectiveForRowsRef.current === id) return;
    loadedObjectiveForRowsRef.current = id;
    let cancelled = false;
    client
      .getObjective(id)
      .then((obj) => {
        if (cancelled || !obj.objectives) return;
        const rows: ObjectiveRowData[] = obj.objectives.map((target, i) => ({
          // Default to Value, not Exact — the ObjectiveSetup dropdown options
          // omit Exact, so a row rehydrated with that type would render with
          // an empty dropdown and silently submit as the placeholder value.
          type: (obj.objectiveTypes?.[i] as ObjectiveType) ?? ObjectiveType.Value,
          target: String(target),
          minBound: String(obj.minimumBounds?.[i] ?? 0),
          maxBound: String(obj.maximumBounds?.[i] ?? 0),
        }));
        setSavedObjectives(rows);
      })
      .catch(() => {
        // Non-fatal — the user can re-enter the form by hand.
      });
    return () => {
      cancelled = true;
    };
  }, [isStateLoaded, client, state.objectiveId]);

  const optimization = useOptimization(client, state.objectiveId, evalConfig);

  /**
   * Recover the input cases for the active project so the user can evaluate and
   * create a *new* trial after a Resume-from-Project landed with no cases (the
   * project DTO omits the large inputCases array). Non-destructive: it only
   * re-reads the project's fixed DOE from the model tree and stores it; the new
   * trial is created later by CaseEvaluation's existing flow, which never
   * modifies prior trials. Returns false when the cases truly aren't available
   * from the API (nothing we can fabricate — the DOE is fixed at creation).
   */
  const reloadProjectInputCases = useCallback(async (): Promise<boolean> => {
    if (!client || !state.modelId || !state.projectId) return false;
    const model = await client.getModel(state.modelId);
    const proj = model.projects?.find((p) => p.id === state.projectId);
    if (proj && Array.isArray(proj.inputCases) && proj.inputCases.length > 0) {
      await updateState({ inputCases: proj.inputCases });
      return true;
    }
    return false;
  }, [client, state.modelId, state.projectId, updateState]);

  const goToStep = useCallback(
    (step: WizardStep) => {
      updateState({ wizardStep: step });
    },
    [updateState]
  );

  /**
   * Centralized hydration for "pick existing" flows. ResumePicker and
   * PickExistingBar both call this. The selection may be partially hydrated —
   * if there's no full `project.trials` available yet (e.g. from the project-
   * mode PickExistingBar), we fetch it here before jumping.
   */
  const handleResume = useCallback(
    async (selection: ResumeSelection, jumpTo: WizardStep) => {
      let project: Project = selection.project;
      // PickExistingBar's project-mode hands us a CatalogProject whose
      // `trials` may be the (possibly stale) embedded list. Refetch the
      // owning model when jumping past EvaluateCases so trial/objective ids
      // are real. The API has no GET /api/projects/{id} — the only way to
      // get the trials[] tree is GET /api/models/{id}, which embeds the
      // full projects[] → trials[] → objectives[] chain.
      const needsHydration =
        jumpTo !== WizardStep.EvaluateCases &&
        (!project.trials || project.trials.length === 0);
      const hydrationModelId =
        (typeof selection.modelId === "number" && selection.modelId > 0
          ? selection.modelId
          : state.modelId) ?? null;
      if (client && needsHydration && hydrationModelId) {
        try {
          const m = await client.getModel(hydrationModelId);
          const fresh = m.projects?.find((p) => p.id === project.id);
          if (fresh) project = fresh;
        } catch {
          // Tolerate hydration failure; we'll proceed with what we have.
        }
      }

      const sameProject = project.id === state.projectId;

      const trial: Trial | undefined =
        selection.trial ??
        (selection.trialId
          ? project.trials.find((t) => t.id === selection.trialId)
          : undefined);

      // Recover user-facing names. The API stores only counts/bounds/types, so
      // names come from in-state (same-project) or the _GMOO_State sheet. The
      // sheet's names are trusted only when it describes *this* project (matching
      // projectId tag, or untagged legacy) and the counts line up — otherwise a
      // different project's labels could leak in. Without this, a resume falls
      // back to "Input N"/"Outcome N" which then show in the template and charts.
      const outputCount = trial?.outputCount ?? 0;
      const savedMatchesProject =
        !!savedStateData &&
        (savedStateData.projectId == null || savedStateData.projectId === project.id);
      const savedVariableNames =
        savedMatchesProject && savedStateData!.variables.length === project.minimums.length
          ? savedStateData!.variables.map((v) => v.name)
          : undefined;
      const savedOutcomeNames =
        savedMatchesProject && savedStateData!.outcomes.length === outputCount
          ? savedStateData!.outcomes.map((o) => o.name)
          : undefined;

      const { inputNames, outcomeNames } = resolveResumeNames({
        inputCount: project.minimums.length,
        outputCount,
        sameProject,
        stateInputNames: state.inputVariables.map((v) => v.name),
        stateOutcomeNames: state.outcomeNames,
        savedVariableNames,
        savedOutcomeNames,
      });

      const inputVariables = project.minimums.map((min, i) => {
        const existing = sameProject ? state.inputVariables[i] : undefined;
        return {
          name: inputNames[i],
          type: project.inputTypes[i] ?? existing?.type ?? "float",
          min,
          max: project.maximums[i] ?? min + 1,
        };
      });

      // Only overwrite modelId when the selection actually carries one
      // (project-mode PickExistingBar / ResumePicker). Trial- and objective-
      // mode in-step switches don't know it; leaving it off preserves the
      // existing state.modelId via the partial-update merge.
      const patch: Partial<WorkbookState> = {
        projectId: project.id,
        trialId: trial?.id ?? null,
        objectiveId: selection.objective?.id ?? null,
        modelName: project.name,
        inputVariables,
        outcomeNames,
        wizardStep: jumpTo,
      };
      if (typeof selection.modelId === "number" && selection.modelId > 0) {
        patch.modelId = selection.modelId;
      }
      // Only overwrite inputCases when the fetched project actually carries
      // them. The listing / embedded DTOs usually omit this large array, and
      // writing `undefined` would blank out the inputCases we already have
      // from a prior createProject — leaving EvaluateCases with nothing to
      // run.
      if (Array.isArray(project.inputCases) && project.inputCases.length > 0) {
        patch.inputCases = project.inputCases;
      }
      // If the user switched to a different project, drop React-only state
      // tied to the old one: cell mappings (evalConfig), the objectives form
      // snapshot, and the example/formulas carry-over. The component-key
      // remount on projectId/trialId handles local component state too.
      if (project.id !== state.projectId) {
        setEvalConfig(null);
        setSavedObjectives(null);
        setSelectedExample(null);
        loadedObjectiveForRowsRef.current = null;
      } else if (trial && trial.id !== state.trialId) {
        // Same project, different trial — objectives form is per-trial.
        setSavedObjectives(null);
        loadedObjectiveForRowsRef.current = null;
      }
      await updateState(patch);
    },
    [
      client,
      updateState,
      state.projectId,
      state.trialId,
      state.inputVariables,
      state.outcomeNames,
      state.modelId,
      savedStateData,
    ]
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
              // Banner is a one-shot signal (e.g. "Activated 'Production'") —
              // clear it when the user moves on so revisiting Connect later
              // doesn't resurface a stale notice.
              setBanner(null);
              goToStep(WizardStep.Resume);
            }}
            banner={banner}
            onDismissBanner={() => setBanner(null)}
          />
        );

      case WizardStep.Resume:
        return (
          <ResumePicker
            client={client}
            aliases={aliases}
            catalog={catalog}
            initialProjectId={state.projectId}
            initialTrialId={state.trialId}
            initialObjectiveId={state.objectiveId}
            onResume={(sel, jumpTo) => void handleResume(sel, jumpTo)}
            onStartFresh={() => goToStep(WizardStep.DefineModel)}
            onBack={() => goToStep(WizardStep.Connect)}
          />
        );

      case WizardStep.DefineModel:
        return (
          <ModelSetup
            client={client}
            headerSlot={
              <PickExistingBar
                mode="project"
                client={client}
                aliases={aliases}
                catalog={catalog}
                currentId={state.projectId}
                onPick={(sel, jumpTo) => void handleResume(sel, jumpTo)}
              />
            }
            initialModelName={state.modelName}
            initialVariables={
              state.inputVariables.length > 0
                ? state.inputVariables
                : savedStateData?.variables.map((v) => ({
                    name: v.name, type: v.type, min: v.min, max: v.max,
                  }))
            }
            onComplete={(data) => {
              setSavedObjectives(null);
              setSelectedExample(data.selectedExample ?? null);
              // Always replace evalConfig: example auto-setup hands us a
              // ready-to-use one, manual flow starts blank. Without the else
              // branch a user who went Back from EvaluateCases and redefined
              // the model would carry the old project's cell mappings into
              // the new project.
              setEvalConfig(data.evalConfig ?? null);

              updateState({
                modelId: data.modelId,
                projectId: data.projectId,
                modelName: data.modelName,
                inputVariables: data.inputVariables,
                // Outcomes are defined on the Evaluate Cases step now; only the
                // example auto-setup path supplies them here. Manual flow seeds
                // the editor empty. Prefer any names already in state (e.g. when
                // re-defining after Back nav) so they aren't blown away.
                outcomeNames: data.outcomeNames ?? state.outcomeNames,
                inputCases: data.inputCases,
                wizardStep: WizardStep.EvaluateCases,
              });
              // Fold the new model/project into the catalog so a later visit
              // to Resume or PickExistingBar's project mode includes it.
              void catalog.refresh();
            }}
            onBack={() => goToStep(WizardStep.Resume)}
          />
        );

      case WizardStep.EvaluateCases:
        return (
          <CaseEvaluation
            // Remount when the project switches so local form state
            // (evalConfig, inputCellMap, outputCellMap, sheetCreated) doesn't
            // carry mappings from the previous project.
            key={state.projectId ?? "none"}
            client={client}
            modelName={state.modelName}
            projectId={state.projectId!}
            variables={state.inputVariables}
            initialOutcomeNames={state.outcomeNames}
            inputCases={state.inputCases ?? []}
            formulas={selectedExample?.setup.formulas}
            initialEvalConfig={evalConfig ?? undefined}
            onReloadInputCases={reloadProjectInputCases}
            headerSlot={
              <PickExistingBar
                mode="trial"
                client={client}
                aliases={aliases}
                modelId={state.modelId}
                projectId={state.projectId}
                currentId={state.trialId}
                onPick={(sel, jumpTo) => void handleResume(sel, jumpTo)}
              />
            }
            onComplete={(trialId, config, outcomeNames) => {
              setEvalConfig(config);
              updateState({
                trialId,
                // Outcomes are finalized on this step now — persist the names
                // the user committed so SetObjectives / Results / Multi-Solve
                // and the charts all use the real labels.
                outcomeNames,
                // Coerce undefined → null so the persisted shape is consistent
                // (EvalConfig.sheetName is undefined in non-contiguous mode;
                // WorkbookState types it as string | null).
                sheetName: config.sheetName ?? null,
                wizardStep: WizardStep.SetObjectives,
              });
              // The new trial isn't in the catalog yet; without this, the
              // PickExistingBar trial dropdown wouldn't see it until refresh.
              void catalog.refresh();
            }}
            onBack={() => goToStep(WizardStep.DefineModel)}
          />
        );

      case WizardStep.SetObjectives:
        return (
          <ObjectiveSetup
            // Remount on trial switch so the objectives form rows reset to
            // the rehydrated savedObjectives for the new trial instead of
            // showing the previous trial's edits.
            key={state.trialId ?? "none"}
            client={client}
            trialId={state.trialId!}
            outcomeNames={state.outcomeNames}
            inputCases={state.inputCases ?? []}
            variables={state.inputVariables}
            evalConfig={evalConfig}
            initialObjectives={savedObjectives ?? undefined}
            exampleObjectives={selectedExample?.objectives}
            headerSlot={
              <PickExistingBar
                mode="objective"
                client={client}
                aliases={aliases}
                modelId={state.modelId}
                projectId={state.projectId}
                trialId={state.trialId}
                currentId={state.objectiveId}
                onPick={(sel, jumpTo) => void handleResume(sel, jumpTo)}
              />
            }
            onComplete={(objectiveId, objectiveRows) => {
              setSavedObjectives(objectiveRows);
              updateState({
                objectiveId,
                wizardStep: WizardStep.Optimize,
              });
              // Same reason as the trial create — keep the cross-step
              // PickExistingBar lists current.
              void catalog.refresh();
            }}
            onBack={() => goToStep(WizardStep.EvaluateCases)}
          />
        );

      case WizardStep.Optimize: {
        // evalConfig is React-only state set during EvaluateCases / example
        // auto-setup. After resume-from-objective the user lands here without
        // it, and clicking Run would silently no-op (see useOptimization
        // guards). Surface why instead.
        //
        // The client check covers reopening a workbook whose
        // activeConnectionId no longer matches any locally-known connection
        // (e.g. opened on a different machine). Without it, Run looks
        // enabled but useOptimization short-circuits on the missing client.
        const notReadyMessage =
          !client
            ? "Your saved API connection isn't available on this machine. Go back to Connect to pick or add one."
            : !state.objectiveId
            ? "No objective selected. Go back and pick one."
            : !evalConfig
            ? "No spreadsheet is bound to this project yet. Go back to Evaluate Cases to bind the input and output ranges, then return here."
            : null;
        return (
          <OptimizationRunner
            state={optimization}
            onRun={(max) => optimization.run(max)}
            onStop={optimization.stop}
            onRunSingle={() => optimization.runSingleIteration()}
            onNext={() => goToStep(WizardStep.Results)}
            onBack={() => goToStep(WizardStep.SetObjectives)}
            notReadyMessage={notReadyMessage}
          />
        );
      }

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
              await catalog.refresh();
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
            GMOO - globalMOO
          </Text>
          <Text className={styles.headerTitle} size={100}>
            Build: {BUILD_TIME}
          </Text>
        </div>
        <WizardStepper currentStep={currentStep} onStepClick={goToStep} />
        <div className={styles.content}>{renderStep()}</div>
      </div>
    </FluentProvider>
  );
};

export default App;
