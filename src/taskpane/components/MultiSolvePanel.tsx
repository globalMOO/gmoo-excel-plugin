import React, { useEffect, useMemo, useState } from "react";
import {
  makeStyles,
  tokens,
  Button,
  Text,
  Input,
  Spinner,
  MessageBar,
  MessageBarBody,
  ProgressBar,
  Badge,
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
} from "@fluentui/react-components";
import {
  Play20Regular,
  Stop20Regular,
  ArrowReset20Regular,
  ArrowLeft20Regular,
  Dismiss20Regular,
  ChartMultiple20Regular,
} from "@fluentui/react-icons";
import type { GmooClient } from "../services/gmooApi";
import type { EvalConfig } from "../services/excelService";
import type { InputVariable } from "../types/workbookState";
import type { ObjectiveRowData } from "./ObjectiveSetup";
import { ObjectiveType } from "../types/gmoo";
import { useMultiSolve } from "../hooks/useMultiSolve";
import { DualRadarCharts, SOLUTION_COLORS } from "./charts/DualRadarCharts";
import { createMultiSolveRadarCharts } from "../services/excelChartService";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    padding: "16px",
  },
  controlsCard: {
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  row: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
  },
  buttonRow: {
    display: "flex",
    gap: "8px",
    flexWrap: "wrap",
  },
  numberInput: {
    width: "70px",
  },
  detailsCard: {
    padding: "10px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
  detailsTitle: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    marginBottom: "6px",
  },
  detailsTitleGrow: {
    flexGrow: 1,
  },
  swatch: {
    width: "12px",
    height: "12px",
    borderRadius: "2px",
  },
  footerRow: {
    display: "flex",
    justifyContent: "space-between",
    marginTop: "8px",
  },
});

interface MultiSolvePanelProps {
  client: GmooClient | null;
  trialId: number | null;
  evalConfig: EvalConfig | null;
  inputVariables: InputVariable[];
  outcomeNames: string[];
  objectiveRows: ObjectiveRowData[] | null;
  onBack: () => void;
}

const NO_TARGET_TYPES = new Set([ObjectiveType.Minimize, ObjectiveType.Maximize]);

export const MultiSolvePanel: React.FC<MultiSolvePanelProps> = ({
  client,
  trialId,
  evalConfig,
  inputVariables,
  outcomeNames,
  objectiveRows,
  onBack,
}) => {
  const styles = useStyles();
  const [numRuns, setNumRuns] = useState(10);
  const [maxIterations, setMaxIterations] = useState(50);
  // Two separate states:
  //   hoveredIdx  — drives chart highlighting; resets to null on mouse leave.
  //   pinnedIdx   — drives the details card; persists so the user can scroll
  //                 the task pane to read without losing the selection.
  const [hoveredIdx, setHoveredIdxRaw] = useState<number | null>(null);
  const [pinnedIdx, setPinnedIdx] = useState<number | null>(null);
  const [exportingCharts, setExportingCharts] = useState(false);
  const [exportError, setExportError] = useState<string | null>(null);
  const [chartsExported, setChartsExported] = useState(false);

  const setHoveredIdx = (idx: number | null) => {
    setHoveredIdxRaw(idx);
    if (idx !== null) setPinnedIdx(idx);
  };

  const multi = useMultiSolve(client, trialId, evalConfig, inputVariables, outcomeNames);

  const canRun =
    client !== null &&
    trialId !== null &&
    evalConfig !== null &&
    inputVariables.length > 0 &&
    objectiveRows !== null &&
    objectiveRows.length > 0 &&
    !multi.isRunning;

  const missingReason =
    objectiveRows === null || objectiveRows.length === 0
      ? "Complete the Set Objectives step before running Multi-Solve."
      : !evalConfig
      ? "No Excel evaluation config — complete the Evaluate Cases step first."
      : !trialId
      ? "No trial available."
      : null;

  const handleExportCharts = async () => {
    setExportingCharts(true);
    setExportError(null);
    setChartsExported(false);
    try {
      await createMultiSolveRadarCharts({
        solutions: multi.solutions,
        inputVariables,
        outcomeNames,
      });
      setChartsExported(true);
    } catch (err) {
      setExportError(err instanceof Error ? err.message : "Failed to export charts.");
    } finally {
      setExportingCharts(false);
    }
  };

  const handleRun = () => {
    if (!objectiveRows) return;
    const targets = objectiveRows.map((o) => parseFloat(o.target) || 0);
    const types = objectiveRows.map((o) => o.type);
    const minBounds = objectiveRows.map((o) => parseFloat(o.minBound) || 0);
    const maxBounds = objectiveRows.map((o) => parseFloat(o.maxBound) || 0);

    multi.run({
      targets,
      types,
      minBounds,
      maxBounds,
      numRuns,
      maxIterations,
    });
  };

  const progressPct = useMemo(() => {
    if (!multi.progress) return 0;
    const { run, totalRuns } = multi.progress;
    return totalRuns === 0 ? 0 : run / totalRuns;
  }, [multi.progress]);

  const stopReasonLabel =
    multi.progress?.stage === "done"
      ? `Completed ${multi.progress.run} of ${multi.progress.totalRuns} runs — ${multi.solutions.length} solution${multi.solutions.length === 1 ? "" : "s"} reported${multi.runsFailed > 0 ? ` (${multi.runsFailed} run${multi.runsFailed === 1 ? "" : "s"} failed)` : ""}. Per-run log on the "GMOO Multi-Solve" sheet — the "Dist. to nearest" column shows normalized-input distance between solutions so you can see how clustered they are.`
      : null;

  // Clear pin when the solution list shrinks (e.g. Clear button) so we don't
  // reference a now-missing index. Also invalidate the "charts exported"
  // indicator — if solutions changed, the previously-exported sheet is stale.
  useEffect(() => {
    if (pinnedIdx !== null && pinnedIdx >= multi.solutions.length) {
      setPinnedIdx(null);
    }
    setChartsExported(false);
  }, [pinnedIdx, multi.solutions.length]);

  const selectedSolution =
    pinnedIdx !== null && multi.solutions[pinnedIdx]
      ? multi.solutions[pinnedIdx]
      : null;

  return (
    <div className={styles.container}>
      <Text weight="semibold" size={400}>
        Multi-Solve
      </Text>
      <Text size={200}>
        Find multiple solutions by running the optimizer from many random
        starting points. Every converged run is reported and overlaid on dual
        radar plots — hover a solution to highlight it in both charts.
      </Text>

      {missingReason && (
        <MessageBar intent="warning">
          <MessageBarBody>{missingReason}</MessageBarBody>
        </MessageBar>
      )}

      {/* Controls */}
      <div className={styles.controlsCard}>
        <div className={styles.row}>
          <Text size={200}>Random starts:</Text>
          <Input
            className={styles.numberInput}
            size="small"
            type="number"
            value={String(numRuns)}
            onChange={(_, data) =>
              setNumRuns(Math.max(1, Math.min(50, parseInt(data.value) || 1)))
            }
            disabled={multi.isRunning}
          />
          <Text size={200}>Max iterations/run:</Text>
          <Input
            className={styles.numberInput}
            size="small"
            type="number"
            value={String(maxIterations)}
            onChange={(_, data) =>
              setMaxIterations(Math.max(1, Math.min(500, parseInt(data.value) || 1)))
            }
            disabled={multi.isRunning}
          />
        </div>
        <div className={styles.buttonRow}>
          {multi.isRunning ? (
            <Button
              icon={<Stop20Regular />}
              appearance="primary"
              onClick={multi.stop}
            >
              Stop
            </Button>
          ) : (
            <Button
              icon={<Play20Regular />}
              appearance="primary"
              onClick={handleRun}
              disabled={!canRun}
            >
              Run Multi-Solve
            </Button>
          )}
          {multi.solutions.length > 0 && !multi.isRunning && (
            <Button
              icon={<ArrowReset20Regular />}
              appearance="secondary"
              onClick={multi.reset}
            >
              Clear
            </Button>
          )}
        </div>
        <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
          Each run creates a new objective on the server ({numRuns} total this
          batch) so it counts against your account's optimization quota.
        </Text>

        {/* Progress */}
        {multi.isRunning && multi.progress && (
          <div style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
            <div className={styles.row}>
              <Spinner size="tiny" />
              <Text size={200}>
                Finding solutions… run {multi.progress.run} of{" "}
                {multi.progress.totalRuns}
                {multi.progress.stage === "iterating" && multi.progress.iteration
                  ? ` — iteration ${multi.progress.iteration}`
                  : ""}
              </Text>
            </div>
            <ProgressBar value={progressPct} />
            <Text size={100}>
              {multi.solutions.length} solution
              {multi.solutions.length === 1 ? "" : "s"}
              {multi.runsFailed > 0 ? ` · ${multi.runsFailed} failed` : ""} ·{" "}
              {multi.runsCompleted}/{multi.progress.totalRuns} runs completed
            </Text>
          </div>
        )}

        {stopReasonLabel && (
          <MessageBar
            intent={multi.solutions.length > 0 ? "success" : "warning"}
          >
            <MessageBarBody>{stopReasonLabel}</MessageBarBody>
          </MessageBar>
        )}

        {multi.error && (
          <MessageBar intent="error">
            <MessageBarBody>{multi.error}</MessageBarBody>
          </MessageBar>
        )}
      </div>

      {/* Radar plots */}
      {multi.solutions.length > 0 && (
        <DualRadarCharts
          solutions={multi.solutions}
          inputVariables={inputVariables}
          outcomeNames={outcomeNames}
          hoveredIdx={hoveredIdx}
          onHoverSolution={setHoveredIdx}
        />
      )}

      {/* Details for pinned solution — persists through scroll so the user
          can read full tables without losing selection on mouse-out. */}
      {selectedSolution && (
        <div className={styles.detailsCard}>
          <div className={styles.detailsTitle}>
            <div
              className={styles.swatch}
              style={{
                backgroundColor:
                  SOLUTION_COLORS[
                    selectedSolution.runIndex % SOLUTION_COLORS.length
                  ],
              }}
            />
            <Text weight="semibold" size={300}>
              Solution {(pinnedIdx ?? 0) + 1}
            </Text>
            <Badge
              appearance="filled"
              color={selectedSolution.satisfied ? "success" : "warning"}
            >
              {selectedSolution.satisfied ? "Satisfied" : "Not satisfied"}
            </Badge>
            <Text size={100} className={styles.detailsTitleGrow}>
              Error: {selectedSolution.l1Norm.toExponential(3)} · iter{" "}
              {selectedSolution.iterations}
            </Text>
            <Button
              size="small"
              appearance="subtle"
              icon={<Dismiss20Regular />}
              onClick={() => setPinnedIdx(null)}
              title="Clear selection"
            />
          </div>

          <Table size="extra-small">
            <TableHeader>
              <TableRow>
                <TableHeaderCell>Input</TableHeaderCell>
                <TableHeaderCell>Value</TableHeaderCell>
              </TableRow>
            </TableHeader>
            <TableBody>
              {inputVariables.map((v, i) => (
                <TableRow key={`in-${i}`}>
                  <TableCell>
                    <Text size={200}>{v.name}</Text>
                  </TableCell>
                  <TableCell>
                    <Text size={200}>
                      {selectedSolution.input?.[i]?.toPrecision(6) ?? "—"}
                    </Text>
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>

          <div style={{ height: "6px" }} />

          <Table size="extra-small">
            <TableHeader>
              <TableRow>
                <TableHeaderCell>Outcome</TableHeaderCell>
                <TableHeaderCell>Target</TableHeaderCell>
                <TableHeaderCell>Achieved</TableHeaderCell>
                <TableHeaderCell>Met?</TableHeaderCell>
              </TableRow>
            </TableHeader>
            <TableBody>
              {outcomeNames.map((name, i) => {
                const res = selectedSolution.results?.[i];
                const isNoTarget =
                  res && NO_TARGET_TYPES.has(res.objectiveType as ObjectiveType);
                return (
                  <TableRow key={`out-${i}`}>
                    <TableCell>
                      <Text size={200}>{name}</Text>
                    </TableCell>
                    <TableCell>
                      <Text size={200}>
                        {res && !isNoTarget
                          ? res.objective.toPrecision(4)
                          : "—"}
                      </Text>
                    </TableCell>
                    <TableCell>
                      <Text size={200}>
                        {selectedSolution.output?.[i]?.toPrecision(4) ?? "—"}
                      </Text>
                    </TableCell>
                    <TableCell>
                      {res ? (
                        <Badge
                          appearance="filled"
                          color={res.satisfied ? "success" : "danger"}
                        >
                          {res.satisfied ? "Yes" : "No"}
                        </Badge>
                      ) : (
                        <Text size={200}>—</Text>
                      )}
                    </TableCell>
                  </TableRow>
                );
              })}
            </TableBody>
          </Table>
        </div>
      )}

      {/* Export radar plots as native Excel charts on a dedicated sheet */}
      {multi.solutions.length > 0 && (
        <div className={styles.buttonRow}>
          <Button
            icon={<ChartMultiple20Regular />}
            appearance="primary"
            onClick={handleExportCharts}
            disabled={exportingCharts}
          >
            {exportingCharts ? (
              <Spinner size="tiny" />
            ) : chartsExported ? (
              "Re-export Radar Charts to Excel"
            ) : (
              "Export Radar Charts to Excel"
            )}
          </Button>
          {chartsExported && !exportError && (
            <Text size={100}>
              Charts created on the "GMOO Multi-Solve Charts" sheet.
            </Text>
          )}
          {exportError && (
            <MessageBar intent="error">
              <MessageBarBody>{exportError}</MessageBarBody>
            </MessageBar>
          )}
        </div>
      )}

      <div className={styles.footerRow}>
        <Button
          icon={<ArrowLeft20Regular />}
          appearance="secondary"
          onClick={onBack}
        >
          Back to Results
        </Button>
      </div>
    </div>
  );
};
