import React from "react";
import {
  makeStyles,
  tokens,
  Text,
} from "@fluentui/react-components";
import {
  CheckmarkCircle20Filled,
  Circle20Regular,
  ArrowCircleRight20Filled,
} from "@fluentui/react-icons";
import { WizardStep, WIZARD_STEP_LABELS } from "../types/workbookState";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "2px",
    padding: "8px 12px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  step: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "4px 0",
    border: "none",
    background: "transparent",
    textAlign: "left",
    width: "100%",
  },
  stepClickable: {
    cursor: "pointer",
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground2Hover,
    },
    ":focus-visible": {
      outline: `2px solid ${tokens.colorBrandStroke1}`,
      outlineOffset: "-2px",
    },
  },
  activeLabel: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorBrandForeground1,
  },
  completedLabel: {
    color: tokens.colorNeutralForeground3,
  },
  pendingLabel: {
    color: tokens.colorNeutralForeground4,
  },
});

interface WizardStepperProps {
  currentStep: WizardStep;
  /** Called when the user clicks a completed step. Pending and the current
   *  step are non-clickable — forward jumps would land on broken state
   *  (e.g. Optimize before there's an objective). */
  onStepClick?: (step: WizardStep) => void;
}

const ALL_STEPS = [
  WizardStep.Connect,
  WizardStep.Resume,
  WizardStep.DefineModel,
  WizardStep.EvaluateCases,
  WizardStep.SetObjectives,
  WizardStep.Optimize,
  WizardStep.Results,
  WizardStep.MultiSolve,
];

export const WizardStepper: React.FC<WizardStepperProps> = ({ currentStep, onStepClick }) => {
  const styles = useStyles();

  return (
    <div className={styles.container}>
      {ALL_STEPS.map((step) => {
        const isCompleted = step < currentStep;
        const isActive = step === currentStep;
        const isClickable = isCompleted && !!onStepClick;

        const className = isClickable
          ? `${styles.step} ${styles.stepClickable}`
          : styles.step;

        const labelClass = isActive
          ? styles.activeLabel
          : isCompleted
            ? styles.completedLabel
            : styles.pendingLabel;

        const icon = isCompleted ? (
          <CheckmarkCircle20Filled primaryFill={tokens.colorPaletteGreenForeground1} />
        ) : isActive ? (
          <ArrowCircleRight20Filled primaryFill={tokens.colorBrandForeground1} />
        ) : (
          <Circle20Regular primaryFill={tokens.colorNeutralForeground4} />
        );

        // Render completed steps as buttons so keyboard + screen-reader users
        // get the same affordance. Pending and current stay as plain divs.
        if (isClickable) {
          return (
            <button
              key={step}
              type="button"
              className={className}
              onClick={() => onStepClick(step)}
              aria-label={`Go back to ${WIZARD_STEP_LABELS[step]}`}
            >
              {icon}
              <Text className={labelClass} size={200}>
                {WIZARD_STEP_LABELS[step]}
              </Text>
            </button>
          );
        }

        return (
          <div key={step} className={styles.step}>
            {icon}
            <Text className={labelClass} size={200}>
              {WIZARD_STEP_LABELS[step]}
            </Text>
          </div>
        );
      })}
    </div>
  );
};
