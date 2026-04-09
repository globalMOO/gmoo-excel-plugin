import React, { useState } from "react";
import {
  makeStyles,
  tokens,
  Button,
  Input,
  Text,
  Spinner,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Card,
  TabList,
  Tab,
} from "@fluentui/react-components";
import { Play20Regular, Pause20Regular, Next20Regular } from "@fluentui/react-icons";
import type { OptimizationState } from "../hooks/useOptimization";
import { filteredL1Norm } from "../types/gmoo";
import { ConvergenceChart } from "./charts/ConvergenceChart";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    padding: "16px",
  },
  controlRow: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
  },
  statsCard: {
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
  statRow: {
    display: "flex",
    justifyContent: "space-between",
    padding: "2px 0",
  },
  buttonRow: {
    display: "flex",
    justifyContent: "space-between",
    marginTop: "12px",
  },
});

interface OptimizationRunnerProps {
  state: OptimizationState;
  onRun: (maxIterations: number) => void;
  onStop: () => void;
  onRunSingle: () => void;
  onNext: () => void;
  onBack: () => void;
}

export const OptimizationRunner: React.FC<OptimizationRunnerProps> = ({
  state,
  onRun,
  onStop,
  onRunSingle,
  onNext,
  onBack,
}) => {
  const styles = useStyles();
  const [maxIterations, setMaxIterations] = useState(100);
  const [activeTab, setActiveTab] = useState<string>("stats");

  const hasResults = state.iterations.length > 0;
  const isDone = !state.isRunning && (state.stopReason !== null || hasResults);
  const isPaused = !state.isRunning && state.stopReason === "Paused by user";

  // Compute filtered error metrics (excluding inequality objectives)
  const filteredBest = hasResults
    ? Math.min(...state.iterations.map((inv) => filteredL1Norm(inv)))
    : null;
  const filteredInitial = state.iterations.length > 0
    ? filteredL1Norm(state.iterations[0])
    : null;
  const filteredCurrent = hasResults
    ? filteredL1Norm(state.iterations[state.iterations.length - 1])
    : null;

  const errorReduction = (filteredInitial !== null && filteredBest !== null && filteredInitial > 0)
    ? ((1 - filteredBest / filteredInitial) * 100)
    : null;

  return (
    <div className={styles.container}>
      <Text weight="semibold" size={400}>
        Optimization
      </Text>

      {/* Auto Mode Controls */}
      <Card>
        <div style={{ padding: "12px" }}>
          <Text weight="semibold" size={300}>
            Auto Mode
          </Text>
          <div className={styles.controlRow} style={{ marginTop: "8px" }}>
            <Text size={200}>Max Iterations:</Text>
            <Input
              size="small"
              type="number"
              value={String(maxIterations)}
              onChange={(_, data) => setMaxIterations(parseInt(data.value) || 100)}
              style={{ width: "80px" }}
              disabled={state.isRunning}
            />
            {!state.isRunning ? (
              <Button
                icon={<Play20Regular />}
                appearance="primary"
                onClick={() => onRun(maxIterations)}
              >
                {isPaused ? "Resume" : "Run"}
              </Button>
            ) : (
              <Button
                icon={<Pause20Regular />}
                appearance="secondary"
                onClick={onStop}
              >
                Pause
              </Button>
            )}
          </div>
        </div>
      </Card>

      {/* Manual Mode Controls */}
      <Card>
        <div style={{ padding: "12px" }}>
          <Text weight="semibold" size={300}>
            Manual Mode
          </Text>
          <div className={styles.controlRow} style={{ marginTop: "8px" }}>
            <Button
              icon={<Next20Regular />}
              appearance="secondary"
              onClick={onRunSingle}
              disabled={state.isRunning}
            >
              Next Iteration
            </Button>
          </div>
        </div>
      </Card>

      {/* Stats / Chart tabs */}
      {(state.isRunning || hasResults) && (
        <>
          <TabList
            selectedValue={activeTab}
            onTabSelect={(_, data) => setActiveTab(data.value as string)}
            size="small"
          >
            <Tab value="stats">Stats</Tab>
            <Tab value="chart">Chart</Tab>
          </TabList>

          {activeTab === "stats" && (
            <div className={styles.statsCard}>
              <div className={styles.statRow}>
                <Text size={200}>Iteration:</Text>
                <Text size={200} weight="semibold">
                  {state.currentIteration}
                  {state.isRunning && <Spinner size="extra-tiny" style={{ marginLeft: "4px" }} />}
                </Text>
              </div>
              <div className={styles.statRow}>
                <Text size={200}>Best Target Error:</Text>
                <Text size={200} weight="semibold">
                  {filteredBest !== null ? filteredBest.toExponential(4) : "—"}
                </Text>
              </div>
              {filteredInitial !== null && (
                <div className={styles.statRow}>
                  <Text size={200}>Initial Target Error:</Text>
                  <Text size={200}>
                    {filteredInitial.toExponential(4)}
                  </Text>
                </div>
              )}
              {filteredCurrent !== null && (
                <div className={styles.statRow}>
                  <Text size={200}>Current Target Error:</Text>
                  <Text size={200}>
                    {filteredCurrent.toExponential(4)}
                  </Text>
                </div>
              )}
              {errorReduction !== null && (
                <div className={styles.statRow}>
                  <Text size={200}>Error Reduction:</Text>
                  <Text size={200} weight="semibold" style={{ color: errorReduction > 99 ? "#107c10" : undefined }}>
                    {errorReduction.toFixed(3)}%
                  </Text>
                </div>
              )}
            </div>
          )}

          {activeTab === "chart" && hasResults && (
            <ConvergenceChart iterations={state.iterations} />
          )}
        </>
      )}

      {/* Stop Reason */}
      {state.stopReason && (
        <MessageBar intent="info">
          <MessageBarBody>
            <MessageBarTitle>Optimization Complete</MessageBarTitle>
            {state.stopReason}
          </MessageBarBody>
        </MessageBar>
      )}

      {/* Error */}
      {state.error && (
        <MessageBar intent="error">
          <MessageBarBody>
            <MessageBarTitle>Error</MessageBarTitle>
            {state.error}
          </MessageBarBody>
        </MessageBar>
      )}

      <div className={styles.buttonRow}>
        <Button appearance="secondary" onClick={onBack} disabled={state.isRunning}>
          Back
        </Button>
        <Button
          appearance="primary"
          onClick={onNext}
          disabled={state.isRunning || !isDone}
        >
          View Results
        </Button>
      </div>
    </div>
  );
};
