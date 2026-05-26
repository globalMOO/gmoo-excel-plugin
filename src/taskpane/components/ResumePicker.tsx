// Wizard step rendered after Connect. Lets the user browse their existing
// models/projects/trials/objectives by name, rename them, pin favorites, and
// jump deep into the wizard — or start a fresh model.
//
// Hierarchy: Model → Project → Trial → Objective. Selecting a project
// triggers a getProject() fetch so we always work with fresh trial +
// objective data instead of the possibly-stale list embedded in /models.

import React, { useCallback, useEffect, useMemo, useState } from "react";
import {
  makeStyles,
  tokens,
  Button,
  Combobox,
  Option,
  Input,
  Text,
  Spinner,
  MessageBar,
  MessageBarBody,
  Card,
} from "@fluentui/react-components";
import {
  Edit20Regular,
  ArrowReset20Regular,
  Checkmark20Regular,
  Dismiss20Regular,
  ArrowSync20Regular,
  Star20Regular,
  Star20Filled,
} from "@fluentui/react-icons";
import type { GmooClient } from "../services/gmooApi";
import type { Project, Trial, Objective, Model } from "../types/gmoo";
import type { EntityKind } from "../types/aliasRegistry";
import type { UseAliasesResult } from "../hooks/useAliases";
import type { UseProjectCatalogResult } from "../hooks/useProjectCatalog";
import { WizardStep } from "../types/workbookState";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "16px",
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
  },
  pickerRow: {
    display: "flex",
    gap: "6px",
    alignItems: "center",
  },
  combobox: {
    flexGrow: 1,
    minWidth: 0,
  },
  renameRow: {
    display: "flex",
    gap: "4px",
    alignItems: "center",
    flexWrap: "wrap",
  },
  renameInput: {
    flexGrow: 1,
  },
  hint: {
    color: tokens.colorNeutralForeground3,
  },
  actions: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    marginTop: "8px",
  },
  buttonRow: {
    display: "flex",
    justifyContent: "space-between",
    marginTop: "12px",
  },
  freshCard: {
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
  },
});

export interface ResumeSelection {
  /**
   * Owning model id. Optional because in-step PickExistingBar trial/objective
   * switches don't know the modelId — only the picker that walks the catalog
   * (project mode + ResumePicker) carries it. When omitted, App.handleResume
   * preserves the existing state.modelId rather than zeroing it.
   */
  modelId?: number;
  projectId: number;
  trialId?: number;
  objectiveId?: number;
  project: Project;
  trial?: Trial;
  objective?: Objective;
}

interface ResumePickerProps {
  client: GmooClient | null;
  aliases: UseAliasesResult;
  catalog: UseProjectCatalogResult;
  initialProjectId: number | null;
  initialTrialId: number | null;
  initialObjectiveId: number | null;
  onResume: (selection: ResumeSelection, jumpTo: WizardStep) => void;
  onStartFresh: () => void;
  onBack: () => void;
}

function trialFallback(t: Trial): string {
  return `Trial #${t.number}`;
}

function objectiveFallback(o: Objective): string {
  return `Objective ${o.id}`;
}

/** Prefix the option text with a star when pinned, so it's identifiable inside the dropdown. */
function withPinMark(text: string, pinned: boolean): string {
  return pinned ? `★ ${text}` : text;
}

export const ResumePicker: React.FC<ResumePickerProps> = ({
  client,
  aliases,
  catalog,
  initialProjectId,
  initialTrialId,
  initialObjectiveId,
  onResume,
  onStartFresh,
  onBack,
}) => {
  const styles = useStyles();

  const [selectedModelId, setSelectedModelId] = useState<number | null>(null);
  const [selectedProjectId, setSelectedProjectId] = useState<number | null>(
    initialProjectId
  );
  const [selectedTrialId, setSelectedTrialId] = useState<number | null>(
    initialTrialId
  );
  const [selectedObjectiveId, setSelectedObjectiveId] = useState<number | null>(
    initialObjectiveId
  );

  // Lazy-fetched model with its embedded projects[]. The /models listing
  // endpoint returns models without their project arrays populated, so we
  // call /models/{id} when the user picks a model to get a fresh project
  // list — and avoid paying for it up front for accounts with many models.
  const [hydratedModel, setHydratedModel] = useState<Model | null>(null);
  const [isLoadingModel, setIsLoadingModel] = useState(false);
  const [modelError, setModelError] = useState<string | null>(null);

  // Inline rename state, keyed by `${kind}:${id}`.
  const [editing, setEditing] = useState<string | null>(null);
  const [editValue, setEditValue] = useState("");

  // Auto-pick the model that owns the initial project, once the catalog loads.
  useEffect(() => {
    if (selectedModelId != null || !initialProjectId) return;
    const owning = catalog.projects.find((p) => p.id === initialProjectId);
    if (owning) setSelectedModelId(owning.modelId);
  }, [catalog.projects, initialProjectId, selectedModelId]);

  // Whenever the selected model changes, fetch its full detail (which carries
  // the projects[] the /models listing omits).
  useEffect(() => {
    if (!client || !selectedModelId) {
      setHydratedModel(null);
      setModelError(null);
      return;
    }
    let cancelled = false;
    setIsLoadingModel(true);
    setModelError(null);
    client
      .getModel(selectedModelId)
      .then((m) => {
        if (!cancelled) setHydratedModel(m);
      })
      .catch((err: unknown) => {
        if (cancelled) return;
        setModelError(err instanceof Error ? err.message : "Failed to load model.");
        setHydratedModel(null);
      })
      .finally(() => {
        if (!cancelled) setIsLoadingModel(false);
      });
    return () => {
      cancelled = true;
    };
  }, [client, selectedModelId]);

  // The hydrated project comes straight from the hydrated model — calling
  // GET /projects/{id} would 404 (route doesn't exist), and the fallback that
  // scans /models[].projects[] returns nothing because /models doesn't embed
  // them. /models/{id} *does* embed projects[] with trials[] populated, so
  // we just pluck the selected one out of the model we already loaded.
  const hydratedProject: Project | null =
    hydratedModel?.projects.find((p) => p.id === selectedProjectId) ?? null;
  const isHydrating = isLoadingModel;
  const hydrateError = modelError;

  // If the previously-selected trial/objective no longer belongs to the
  // hydrated project, clear it.
  useEffect(() => {
    if (!hydratedProject) return;
    if (
      selectedTrialId &&
      !hydratedProject.trials.some((t) => t.id === selectedTrialId)
    ) {
      setSelectedTrialId(null);
      setSelectedObjectiveId(null);
    }
  }, [hydratedProject, selectedTrialId]);

  // --- Derived option lists, sorted pinned-first ---

  const sortedModels: Model[] = useMemo(
    () =>
      aliases.sortByPinned(
        catalog.models,
        () => "model",
        (m) => m.id,
        (m) => aliases.getName("model", m.id, m.name)
      ),
    [aliases, catalog.models]
  );

  const projectsInSelectedModel: Project[] = useMemo(() => {
    const list = hydratedModel?.projects ?? [];
    if (list.length === 0) return [];
    return aliases.sortByPinned(
      list,
      () => "project",
      (p) => p.id,
      (p) => aliases.getName("project", p.id, p.name)
    );
  }, [aliases, hydratedModel]);

  const sortedTrials: Trial[] = useMemo(() => {
    if (!hydratedProject) return [];
    return aliases.sortByPinned(
      hydratedProject.trials,
      () => "trial",
      (t) => t.id,
      (t) => aliases.getName("trial", t.id, trialFallback(t))
    );
  }, [aliases, hydratedProject]);

  const selectedModel = catalog.models.find((m) => m.id === selectedModelId) ?? null;
  const selectedProject =
    hydratedModel?.projects.find((p) => p.id === selectedProjectId) ?? null;
  const selectedTrial =
    hydratedProject?.trials.find((t) => t.id === selectedTrialId) ?? null;
  const selectedObjective =
    selectedTrial?.objectives.find((o) => o.id === selectedObjectiveId) ?? null;

  const sortedObjectives: Objective[] = useMemo(() => {
    if (!selectedTrial) return [];
    return aliases.sortByPinned(
      selectedTrial.objectives,
      () => "objective",
      (o) => o.id,
      (o) => aliases.getName("objective", o.id, objectiveFallback(o))
    );
  }, [aliases, selectedTrial]);

  // --- Rename helpers ---

  const startEditing = (kind: EntityKind, id: number, current: string) => {
    setEditing(`${kind}:${id}`);
    setEditValue(current);
  };
  const cancelEditing = () => {
    setEditing(null);
    setEditValue("");
  };
  const saveEditing = useCallback(
    async (kind: EntityKind, id: number) => {
      await aliases.rename(kind, id, editValue);
      setEditing(null);
      setEditValue("");
    },
    [aliases, editValue]
  );

  const renderRenameRow = (
    kind: EntityKind,
    id: number,
    currentLabel: string,
    fallback: string
  ) => {
    const isEditing = editing === `${kind}:${id}`;
    const hasAlias = currentLabel !== fallback;
    const pinned = aliases.isPinned(kind, id);
    return (
      <div className={styles.renameRow}>
        {isEditing ? (
          <>
            <Input
              className={styles.renameInput}
              size="small"
              value={editValue}
              onChange={(_, data) => setEditValue(data.value)}
              placeholder={fallback}
              autoFocus
              onKeyDown={(e) => {
                if (e.key === "Enter") void saveEditing(kind, id);
                if (e.key === "Escape") cancelEditing();
              }}
            />
            <Button
              icon={<Checkmark20Regular />}
              size="small"
              appearance="subtle"
              title="Save"
              onClick={() => void saveEditing(kind, id)}
            />
            <Button
              icon={<Dismiss20Regular />}
              size="small"
              appearance="subtle"
              title="Cancel"
              onClick={cancelEditing}
            />
          </>
        ) : (
          <>
            <Button
              icon={pinned ? <Star20Filled /> : <Star20Regular />}
              size="small"
              appearance="subtle"
              title={pinned ? "Unpin" : "Pin to top"}
              onClick={() => void aliases.togglePin(kind, id)}
            >
              {pinned ? "Pinned" : "Pin"}
            </Button>
            <Button
              icon={<Edit20Regular />}
              size="small"
              appearance="subtle"
              title="Rename"
              onClick={() => startEditing(kind, id, hasAlias ? currentLabel : "")}
            >
              Rename
            </Button>
            {hasAlias && (
              <Button
                icon={<ArrowReset20Regular />}
                size="small"
                appearance="subtle"
                title="Reset to default name"
                onClick={() => void aliases.reset(kind, id)}
              >
                Reset
              </Button>
            )}
          </>
        )}
      </div>
    );
  };

  // --- Resume handlers ---

  const handleResume = (jumpTo: WizardStep) => {
    if (!hydratedProject || !selectedProject || !selectedModelId) return;
    const sel: ResumeSelection = {
      // Use selectedModelId (the dropdown selection) rather than reading it
      // off the project — the bare Project DTO from /models/{id} doesn't
      // carry modelId since the model is implicit in the request URL.
      modelId: selectedModelId,
      projectId: selectedProject.id,
      project: hydratedProject,
      trialId: selectedTrial?.id,
      trial: selectedTrial ?? undefined,
      objectiveId: selectedObjective?.id,
      objective: selectedObjective ?? undefined,
    };
    onResume(sel, jumpTo);
  };

  // --- Rendering ---

  if (catalog.isLoading && catalog.models.length === 0) {
    return (
      <div className={styles.container}>
        <Spinner label="Loading models…" />
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <Text weight="semibold" size={400}>
        Resume an existing project, or start a new one
      </Text>

      {catalog.error && (
        <MessageBar intent="error">
          <MessageBarBody>{catalog.error}</MessageBarBody>
        </MessageBar>
      )}

      <Card className={styles.freshCard}>
        <Text size={200} className={styles.hint}>
          Want to train a new project from scratch?
        </Text>
        <div style={{ marginTop: "8px" }}>
          <Button appearance="primary" onClick={onStartFresh}>
            Start a new model
          </Button>
        </div>
      </Card>

      {/* --- Model picker --- */}
      <div className={styles.section}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <Text weight="semibold" size={300}>
            Model
          </Text>
          <Button
            icon={<ArrowSync20Regular />}
            size="small"
            appearance="subtle"
            onClick={() => void catalog.refresh()}
            disabled={catalog.isLoading}
            title="Reload catalog"
          >
            Refresh
          </Button>
        </div>
        <div className={styles.pickerRow}>
          <Combobox
            className={styles.combobox}
            placeholder={
              catalog.models.length === 0
                ? "No models found for this connection"
                : "Pick a model…"
            }
            value={
              selectedModel
                ? aliases.getName("model", selectedModel.id, selectedModel.name)
                : ""
            }
            selectedOptions={selectedModelId ? [String(selectedModelId)] : []}
            onOptionSelect={(_, data) => {
              const id = data.optionValue ? Number(data.optionValue) : null;
              setSelectedModelId(id);
              setSelectedProjectId(null);
              setSelectedTrialId(null);
              setSelectedObjectiveId(null);
            }}
            disabled={catalog.models.length === 0}
          >
            {sortedModels.map((m) => {
              const alias = aliases.getName("model", m.id, m.name);
              const pinned = aliases.isPinned("model", m.id);
              const text = withPinMark(alias, pinned);
              return (
                <Option key={m.id} value={String(m.id)} text={text}>
                  {text}
                </Option>
              );
            })}
          </Combobox>
        </div>
        {selectedModel &&
          renderRenameRow(
            "model",
            selectedModel.id,
            aliases.getName("model", selectedModel.id, selectedModel.name),
            selectedModel.name
          )}
      </div>

      {/* --- Project picker (lazy-loaded from getModel(selectedModelId)) --- */}
      {selectedModelId && (
        <div className={styles.section}>
          <Text weight="semibold" size={300}>
            Project
          </Text>
          {isLoadingModel && <Spinner size="tiny" label="Loading projects…" />}
          {modelError && (
            <MessageBar intent="error">
              <MessageBarBody>{modelError}</MessageBarBody>
            </MessageBar>
          )}
          <div className={styles.pickerRow}>
            <Combobox
              className={styles.combobox}
              placeholder={
                isLoadingModel
                  ? "Loading…"
                  : projectsInSelectedModel.length === 0
                  ? "This model has no projects yet"
                  : "Pick a project…"
              }
              value={
                selectedProject
                  ? aliases.getName("project", selectedProject.id, selectedProject.name)
                  : ""
              }
              selectedOptions={selectedProjectId ? [String(selectedProjectId)] : []}
              onOptionSelect={(_, data) => {
                const id = data.optionValue ? Number(data.optionValue) : null;
                setSelectedProjectId(id);
                setSelectedTrialId(null);
                setSelectedObjectiveId(null);
              }}
              disabled={projectsInSelectedModel.length === 0}
            >
              {projectsInSelectedModel.map((p) => {
                const alias = aliases.getName("project", p.id, p.name);
                const pinned = aliases.isPinned("project", p.id);
                const text = withPinMark(alias, pinned);
                return (
                  <Option key={p.id} value={String(p.id)} text={text}>
                    {text}
                  </Option>
                );
              })}
            </Combobox>
          </div>
          {selectedProject &&
            renderRenameRow(
              "project",
              selectedProject.id,
              aliases.getName("project", selectedProject.id, selectedProject.name),
              selectedProject.name
            )}
        </div>
      )}

      {/* --- Trial picker --- */}
      {selectedProjectId && (
        <div className={styles.section}>
          <Text weight="semibold" size={300}>
            Trial
          </Text>
          {isHydrating ? (
            <Spinner size="tiny" label="Loading trials…" />
          ) : hydrateError ? (
            <MessageBar intent="error">
              <MessageBarBody>{hydrateError}</MessageBarBody>
            </MessageBar>
          ) : hydratedProject && hydratedProject.trials.length === 0 ? (
            <Text size={200} className={styles.hint}>
              This project has no trials yet. Resume from Project to load output cases.
            </Text>
          ) : hydratedProject ? (
            <>
              <div className={styles.pickerRow}>
                <Combobox
                  className={styles.combobox}
                  placeholder="Pick a trial…"
                  value={
                    selectedTrial
                      ? aliases.getName("trial", selectedTrial.id, trialFallback(selectedTrial))
                      : ""
                  }
                  selectedOptions={selectedTrialId ? [String(selectedTrialId)] : []}
                  onOptionSelect={(_, data) => {
                    const id = data.optionValue ? Number(data.optionValue) : null;
                    setSelectedTrialId(id);
                    setSelectedObjectiveId(null);
                  }}
                >
                  {sortedTrials.map((t) => {
                    const alias = aliases.getName("trial", t.id, trialFallback(t));
                    const pinned = aliases.isPinned("trial", t.id);
                    const text = withPinMark(alias, pinned);
                    return (
                      <Option key={t.id} value={String(t.id)} text={text}>
                        {text}
                      </Option>
                    );
                  })}
                </Combobox>
              </div>
              {selectedTrial &&
                renderRenameRow(
                  "trial",
                  selectedTrial.id,
                  aliases.getName("trial", selectedTrial.id, trialFallback(selectedTrial)),
                  trialFallback(selectedTrial)
                )}
            </>
          ) : null}
        </div>
      )}

      {/* --- Objective picker --- */}
      {selectedTrialId && selectedTrial && (
        <div className={styles.section}>
          <Text weight="semibold" size={300}>
            Objective
          </Text>
          {selectedTrial.objectives.length === 0 ? (
            <Text size={200} className={styles.hint}>
              This trial has no objectives yet. Resume from Trial to set them.
            </Text>
          ) : (
            <>
              <div className={styles.pickerRow}>
                <Combobox
                  className={styles.combobox}
                  placeholder="Pick an objective…"
                  value={
                    selectedObjective
                      ? aliases.getName(
                          "objective",
                          selectedObjective.id,
                          objectiveFallback(selectedObjective)
                        )
                      : ""
                  }
                  selectedOptions={
                    selectedObjectiveId ? [String(selectedObjectiveId)] : []
                  }
                  onOptionSelect={(_, data) => {
                    const id = data.optionValue ? Number(data.optionValue) : null;
                    setSelectedObjectiveId(id);
                  }}
                >
                  {sortedObjectives.map((o) => {
                    const alias = aliases.getName("objective", o.id, objectiveFallback(o));
                    const pinned = aliases.isPinned("objective", o.id);
                    const text = withPinMark(alias, pinned);
                    return (
                      <Option key={o.id} value={String(o.id)} text={text}>
                        {text}
                      </Option>
                    );
                  })}
                </Combobox>
              </div>
              {selectedObjective &&
                renderRenameRow(
                  "objective",
                  selectedObjective.id,
                  aliases.getName(
                    "objective",
                    selectedObjective.id,
                    objectiveFallback(selectedObjective)
                  ),
                  objectiveFallback(selectedObjective)
                )}
            </>
          )}
        </div>
      )}

      {/* --- Resume actions ---
          The Project-level button creates a *new* trial on submit, so when the
          project already has trials we relabel it to make that side-effect
          explicit. (Picking an existing trial above re-uses it instead.) */}
      <div className={styles.actions}>
        <Button
          appearance="primary"
          disabled={!selectedObjective || isHydrating}
          onClick={() => handleResume(WizardStep.Optimize)}
        >
          Resume from Objective → Optimize
        </Button>
        <Button
          appearance="secondary"
          disabled={!selectedTrial || isHydrating}
          onClick={() => handleResume(WizardStep.SetObjectives)}
        >
          Resume from Trial → Set Objectives
        </Button>
        <Button
          appearance="secondary"
          disabled={!hydratedProject || isHydrating}
          onClick={() => handleResume(WizardStep.EvaluateCases)}
        >
          {hydratedProject && hydratedProject.trials.length > 0
            ? "Add new trial → Evaluate Cases"
            : "Resume from Project → Evaluate Cases"}
        </Button>
      </div>

      <div className={styles.buttonRow}>
        <Button appearance="secondary" onClick={onBack}>
          Back
        </Button>
        <span />
      </div>
    </div>
  );
};
