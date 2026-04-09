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
}

const ALL_STEPS = [
  WizardStep.Connect,
  WizardStep.DefineModel,
  WizardStep.EvaluateCases,
  WizardStep.SetObjectives,
  WizardStep.Optimize,
  WizardStep.Results,
];

export const WizardStepper: React.FC<WizardStepperProps> = ({ currentStep }) => {
  const styles = useStyles();

  return (
    <div className={styles.container}>
      {ALL_STEPS.map((step) => {
        const isCompleted = step < currentStep;
        const isActive = step === currentStep;

        return (
          <div key={step} className={styles.step}>
            {isCompleted ? (
              <CheckmarkCircle20Filled primaryFill={tokens.colorPaletteGreenForeground1} />
            ) : isActive ? (
              <ArrowCircleRight20Filled primaryFill={tokens.colorBrandForeground1} />
            ) : (
              <Circle20Regular primaryFill={tokens.colorNeutralForeground4} />
            )}
            <Text
              className={
                isActive
                  ? styles.activeLabel
                  : isCompleted
                    ? styles.completedLabel
                    : styles.pendingLabel
              }
              size={200}
            >
              {WIZARD_STEP_LABELS[step]}
            </Text>
          </div>
        );
      })}
    </div>
  );
};
