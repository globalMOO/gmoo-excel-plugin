// Compact selector rendered at the top of ModelSetup / CaseEvaluation /
// ObjectiveSetup. Lets the user swap to a different existing entity without
// leaving the current step. Selecting an item dispatches the same onResume
// callback used by ResumePicker, so all state-hydration logic lives in
// App.tsx.

import React, { useEffect, useState } from "react";
import {
  makeStyles,
  tokens,
  Combobox,
  Option,
  Text,
  Spinner,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import { GmooApiError, type GmooClient } from "../services/gmooApi";
import type { Project, Trial, Objective } from "../types/gmoo";
import type { UseAliasesResult } from "../hooks/useAliases";
import type { UseProjectCatalogResult } from "../hooks/useProjectCatalog";
import type { ResumeSelection } from "./ResumePicker";
import { WizardStep } from "../types/workbookState";
import type { EntityKind } from "../types/aliasRegistry";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    padding: "8px 12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  row: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
  },
  label: {
    color: tokens.colorNeutralForeground3,
    whiteSpace: "nowrap",
  },
  combobox: {
    flexGrow: 1,
    minWidth: 0,
  },
});

type Mode = "project" | "trial" | "objective";

interface CommonProps {
  client: GmooClient | null;
  aliases: UseAliasesResult;
  currentId: number | null;
  onPick: (selection: ResumeSelection, jumpTo: WizardStep) => void;
}

interface ProjectModeProps extends CommonProps {
  mode: "project";
  catalog: UseProjectCatalogResult;
}

interface TrialModeProps extends CommonProps {
  mode: "trial";
  /** Used to hydrate trial list via getModel(modelId), avoiding the missing
   *  GET /projects/{id} route and the often-stale all-models listing. */
  modelId: number | null;
  projectId: number | null;
}

interface ObjectiveModeProps extends CommonProps {
  mode: "objective";
  modelId: number | null;
  projectId: number | null;
  trialId: number | null;
}

export type PickExistingBarProps = ProjectModeProps | TrialModeProps | ObjectiveModeProps;

export const PickExistingBar: React.FC<PickExistingBarProps> = (props) => {
  const styles = useStyles();
  const { client, aliases, currentId, onPick, mode } = props;

  // For trial/objective modes we lazily hydrate the project.
  const [hydratedProject, setHydratedProject] = useState<Project | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const needsHydration = mode === "trial" || mode === "objective";
  const projectIdForHydration = needsHydration
    ? (props as TrialModeProps | ObjectiveModeProps).projectId
    : null;
  const modelIdForHydration = needsHydration
    ? (props as TrialModeProps | ObjectiveModeProps).modelId
    : null;

  useEffect(() => {
    if (!needsHydration) {
      setHydratedProject(null);
      return;
    }
    if (!client || !projectIdForHydration || !modelIdForHydration) {
      // Without a modelId we can't reach the project tree — the API has no
      // GET /api/projects/{id}, only GET /api/models/{id} (which embeds
      // projects[].trials[].objectives[]).
      setHydratedProject(null);
      return;
    }
    let cancelled = false;
    setIsLoading(true);
    setError(null);
    client
      .getModel(modelIdForHydration)
      .then((m) => {
        if (cancelled) return;
        const p = m.projects?.find((pp) => pp.id === projectIdForHydration);
        setHydratedProject(p ?? null);
      })
      .catch((err: unknown) => {
        if (cancelled) return;
        setHydratedProject(null);
        // 404 on the model itself is silently absorbed — likely a freshly-
        // created model the listing hasn't caught up to yet. The placeholder
        // copy ("No trials in this project") is the right user-facing signal.
        if (err instanceof GmooApiError && err.status === 404) return;
        setError(err instanceof Error ? err.message : "Failed to load project.");
      })
      .finally(() => {
        if (!cancelled) setIsLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [client, projectIdForHydration, modelIdForHydration, needsHydration]);

  // --- Build the option list and resume jump target for the current mode ---

  interface PickOption {
    id: number;
    text: string;
    kind: EntityKind;
  }

  let label: string;
  let placeholder: string;
  let rawOptions: PickOption[] = [];
  let resolve: ((id: number) => ResumeSelection | null) | null = null;
  let jumpTo: WizardStep;

  if (mode === "project") {
    // Filter out the currently-active project — picking it again would do
    // nothing useful, and listing it suggests it's a meaningful option.
    const projects = (props as ProjectModeProps).catalog.projects.filter(
      (p) => p.id !== currentId
    );
    label = "Switch project";
    placeholder = projects.length === 0 ? "No other projects available" : "Pick a project…";
    jumpTo = WizardStep.EvaluateCases;
    rawOptions = projects.map((p) => ({
      id: p.id,
      kind: "project" as const,
      text: `${aliases.getName("project", p.id, p.name)} — ${p.modelName}`,
    }));
    resolve = (id) => {
      const p = projects.find((pp) => pp.id === id);
      if (!p) return null;
      return {
        modelId: p.modelId,
        projectId: p.id,
        // The catalog entry already includes trials[] (the catalog hook
        // fans out getModel for every model to populate them). App.handleResume
        // re-fetches via getModel(modelId) when jumping past EvaluateCases,
        // so stale trial lists are caught there.
        project: p,
      };
    };
  } else if (mode === "trial") {
    label = "Switch trial";
    jumpTo = WizardStep.SetObjectives;
    const trials = hydratedProject?.trials ?? [];
    placeholder = trials.length === 0 ? "No trials in this project" : "Pick a trial…";
    rawOptions = trials.map((t: Trial) => ({
      id: t.id,
      kind: "trial" as const,
      text: aliases.getName("trial", t.id, `Trial #${t.number}`),
    }));
    resolve = (id) => {
      if (!hydratedProject) return null;
      const t = hydratedProject.trials.find((tt) => tt.id === id);
      if (!t) return null;
      // modelId omitted on purpose — App.handleResume preserves the existing
      // state.modelId. Don't pass 0, that gets written verbatim and breaks
      // the next render's getModel(0) call.
      return {
        projectId: hydratedProject.id,
        project: hydratedProject,
        trialId: t.id,
        trial: t,
      };
    };
  } else {
    label = "Switch objective";
    jumpTo = WizardStep.Optimize;
    const trial = hydratedProject?.trials.find(
      (t) => t.id === (props as ObjectiveModeProps).trialId
    );
    const objs = trial?.objectives ?? [];
    placeholder =
      objs.length === 0 ? "No objectives in this trial" : "Pick an objective…";
    rawOptions = objs.map((o: Objective) => ({
      id: o.id,
      kind: "objective" as const,
      text: aliases.getName("objective", o.id, `Objective ${o.id}`),
    }));
    resolve = (id) => {
      if (!hydratedProject || !trial) return null;
      const o = trial.objectives.find((oo) => oo.id === id);
      if (!o) return null;
      // See trial-mode comment above re: omitting modelId.
      return {
        projectId: hydratedProject.id,
        project: hydratedProject,
        trialId: trial.id,
        trial,
        objectiveId: o.id,
        objective: o,
      };
    };
  }

  // Pinned-first, then alphabetical. Pinned items are also prefixed with "★ "
  // so the badge is visible even after Combobox sort/scroll.
  const sortedOptions = aliases.sortByPinned(
    rawOptions,
    (o) => o.kind,
    (o) => o.id,
    (o) => o.text
  );
  const options = sortedOptions.map((o) => ({
    ...o,
    text: aliases.isPinned(o.kind, o.id) ? `★ ${o.text}` : o.text,
  }));

  const currentOption = options.find((o) => o.id === currentId);
  const currentText = currentOption?.text ?? "";

  return (
    <div className={styles.container}>
      <div className={styles.row}>
        <Text size={200} className={styles.label}>
          {label}:
        </Text>
        {isLoading ? (
          <Spinner size="tiny" />
        ) : (
          <Combobox
            className={styles.combobox}
            placeholder={placeholder}
            value={currentText}
            selectedOptions={currentOption ? [String(currentOption.id)] : []}
            onOptionSelect={(_, data) => {
              const id = data.optionValue ? Number(data.optionValue) : null;
              if (id == null) return;
              const sel = resolve?.(id);
              if (sel) onPick(sel, jumpTo);
            }}
            disabled={options.length === 0}
            size="small"
          >
            {options.map((o) => (
              <Option key={o.id} value={String(o.id)} text={o.text}>
                {o.text}
              </Option>
            ))}
          </Combobox>
        )}
      </div>
      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}
    </div>
  );
};
