import React, { useState, useEffect, useRef } from "react";
import {
  makeStyles,
  tokens,
  Button,
  Text,
  Spinner,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  Badge,
} from "@fluentui/react-components";
import { ChartMultiple20Regular, ArrowReset20Regular } from "@fluentui/react-icons";
import type { Inverse } from "../types/gmoo";
import { getStopReason, getStopReasonLabel, StopReason, filteredL1Norm, isTargetBasedType } from "../types/gmoo";
import { createResultsCharts } from "../services/excelChartService";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    padding: "16px",
  },
  bestCaseCard: {
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
  statRow: {
    display: "flex",
    justifyContent: "space-between",
    padding: "4px 0",
  },
  buttonRow: {
    display: "flex",
    gap: "8px",
    marginTop: "12px",
  },
});

interface ResultsSummaryProps {
  iterations: Inverse[];
  inputVariableNames: string[];
  outcomeNames: string[];
  onStartOver: () => void;
}

export const ResultsSummary: React.FC<ResultsSummaryProps> = ({
  iterations,
  inputVariableNames,
  outcomeNames,
  onStartOver,
}) => {
  const styles = useStyles();
  const [isCreatingCharts, setIsCreatingCharts] = useState(false);
  const [chartsCreated, setChartsCreated] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const bottomRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, []);

  if (iterations.length === 0) {
    return (
      <div className={styles.container}>
        <Text>No optimization results available.</Text>
        <Button onClick={onStartOver}>Start Over</Button>
      </div>
    );
  }

  // Find best iteration (lowest filtered target error)
  const bestInverse = iterations.reduce((best, inv) =>
    filteredL1Norm(inv) < filteredL1Norm(best) ? inv : best
  );

  const lastInverse = iterations[iterations.length - 1];
  const stopReason = getStopReason(lastInverse);

  const handleCreateCharts = async () => {
    setIsCreatingCharts(true);
    setError(null);
    try {
      await createResultsCharts({
        iterations,
        inputVariableNames,
        outcomeNames,
      });
      setChartsCreated(true);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to create charts.");
    } finally {
      setIsCreatingCharts(false);
    }
  };

  return (
    <div className={styles.container}>
      <Text weight="semibold" size={400}>
        Optimization Results
      </Text>

      {/* Stop Reason */}
      <MessageBar
        intent={stopReason === StopReason.Satisfied ? "success" : "warning"}
      >
        <MessageBarBody>
          <MessageBarTitle>
            {stopReason === StopReason.Satisfied ? "Objective Satisfied" : "Optimization Complete"}
          </MessageBarTitle>
          {getStopReasonLabel(stopReason)} after {iterations.length} iterations.
        </MessageBarBody>
      </MessageBar>

      {/* Best Case Summary */}
      <div className={styles.bestCaseCard}>
        <Text weight="semibold" size={300}>
          Best Case (Iteration {bestInverse.iteration})
        </Text>
        <div className={styles.statRow}>
          <Text size={200}>Target Error:</Text>
          <Text size={200} weight="semibold">
            {filteredL1Norm(bestInverse).toExponential(6)}
          </Text>
        </div>
      </div>

      {/* Optimal Input Values */}
      <Text weight="semibold" size={300}>
        Optimal Inputs
      </Text>
      <Table size="extra-small">
        <TableHeader>
          <TableRow>
            <TableHeaderCell>Variable</TableHeaderCell>
            <TableHeaderCell>Value</TableHeaderCell>
          </TableRow>
        </TableHeader>
        <TableBody>
          {inputVariableNames.map((name, i) => (
            <TableRow key={i}>
              <TableCell>
                <Text size={200}>{name}</Text>
              </TableCell>
              <TableCell>
                <Text size={200}>
                  {bestInverse.input?.[i]?.toPrecision(6) ?? "—"}
                </Text>
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>

      {/* Per-Outcome Breakdown */}
      {bestInverse.results && bestInverse.results.length > 0 && (
        <>
          <Text weight="semibold" size={300}>
            Outcome Detail
          </Text>
          <Table size="extra-small">
            <TableHeader>
              <TableRow>
                <TableHeaderCell>Outcome</TableHeaderCell>
                <TableHeaderCell>Target</TableHeaderCell>
                <TableHeaderCell>Achieved</TableHeaderCell>
                <TableHeaderCell>Error</TableHeaderCell>
                <TableHeaderCell>Met?</TableHeaderCell>
              </TableRow>
            </TableHeader>
            <TableBody>
              {bestInverse.results.map((result, i) => (
                <TableRow key={i}>
                  <TableCell>
                    <Text size={200}>{outcomeNames[i] ?? `Outcome ${i + 1}`}</Text>
                  </TableCell>
                  <TableCell>
                    <Text size={200}>{result.objective.toPrecision(4)}</Text>
                  </TableCell>
                  <TableCell>
                    <Text size={200}>{result.output.toPrecision(4)}</Text>
                  </TableCell>
                  <TableCell>
                    <Text size={200}>
                      {isTargetBasedType(result.objectiveType) ? result.error.toExponential(2) : "—"}
                    </Text>
                  </TableCell>
                  <TableCell>
                    <Badge
                      appearance="filled"
                      color={result.satisfied ? "success" : "danger"}
                    >
                      {result.satisfied ? "Yes" : "No"}
                    </Badge>
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </>
      )}

      {/* Actions */}
      <div className={styles.buttonRow}>
        <Button
          icon={<ChartMultiple20Regular />}
          appearance="primary"
          onClick={handleCreateCharts}
          disabled={isCreatingCharts || chartsCreated}
        >
          {isCreatingCharts ? (
            <Spinner size="tiny" />
          ) : chartsCreated ? (
            "Charts Created"
          ) : (
            "Generate Excel Charts"
          )}
        </Button>
        <Button
          icon={<ArrowReset20Regular />}
          appearance="secondary"
          onClick={onStartOver}
        >
          Start Over
        </Button>
      </div>

      {chartsCreated && (
        <MessageBar intent="success">
          <MessageBarBody>
            Charts created on the "VSME Results" sheet.
          </MessageBarBody>
        </MessageBar>
      )}

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}
      <div ref={bottomRef} />
    </div>
  );
};
