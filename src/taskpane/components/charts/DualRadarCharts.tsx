// Dual radar plots (inputs + outputs) with complementary highlighting.
// Hovering a solution polygon (or legend chip) in either chart highlights that
// solution in BOTH charts and dims the others. Mirrors the Multi-Solve radar
// behavior in the /decisions repo, ported to Chart.js + Fluent UI.
import React, { useCallback, useMemo, useRef } from "react";
import { Radar } from "react-chartjs-2";
import {
  Chart as ChartJS,
  RadialLinearScale,
  PointElement,
  LineElement,
  Filler,
  Tooltip,
  Legend,
  RadarController,
} from "chart.js";
import { makeStyles, tokens, Text } from "@fluentui/react-components";
import type { MultiSolveSolution } from "../../hooks/useMultiSolve";
import type { InputVariable } from "../../types/workbookState";

ChartJS.register(
  RadialLinearScale,
  PointElement,
  LineElement,
  Filler,
  Tooltip,
  Legend,
  RadarController
);

// Solution color palette — matches the /decisions repo
export const SOLUTION_COLORS = [
  "#e94560", "#3498db", "#27ae60", "#f39c12", "#9b59b6",
  "#1abc9c", "#e67e22", "#e74c3c", "#2ecc71", "#8e44ad",
  "#f1c40f", "#16a085", "#d35400", "#c0392b", "#2980b9",
];

interface DualRadarChartsProps {
  solutions: MultiSolveSolution[];
  inputVariables: InputVariable[];
  outcomeNames: string[];
  hoveredIdx: number | null;
  onHoverSolution: (idx: number | null) => void;
}

const useStyles = makeStyles({
  wrapper: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  chartRow: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
  },
  chartCard: {
    padding: "8px",
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  chartTitle: {
    marginBottom: "4px",
    display: "block",
    textAlign: "center",
  },
  chartBox: {
    height: "260px",
    position: "relative",
  },
  legendRow: {
    display: "flex",
    flexWrap: "wrap",
    gap: "6px",
    padding: "8px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
  legendChip: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    padding: "4px 10px",
    borderRadius: tokens.borderRadiusCircular,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    cursor: "pointer",
    transition: "opacity 0.15s, background-color 0.15s",
    userSelect: "none",
  },
  legendSwatch: {
    width: "12px",
    height: "12px",
    borderRadius: "2px",
    flexShrink: 0,
  },
});

function hexToRgba(hex: string, alpha: number): string {
  const clean = hex.replace("#", "");
  const r = parseInt(clean.substring(0, 2), 16);
  const g = parseInt(clean.substring(2, 4), 16);
  const b = parseInt(clean.substring(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

// Alpha rules match the /decisions repo highlighting:
//   normal    -> 0.15 fill
//   hovered   -> 0.40 fill
//   dimmed    -> 0.05 fill
function getFillAlpha(hoveredIdx: number | null, i: number): number {
  if (hoveredIdx === null) return 0.15;
  if (hoveredIdx === i) return 0.4;
  return 0.05;
}

function getBorderAlpha(hoveredIdx: number | null, i: number): number {
  if (hoveredIdx === null) return 0.9;
  if (hoveredIdx === i) return 1.0;
  return 0.3;
}

function getBorderWidth(hoveredIdx: number | null, i: number): number {
  if (hoveredIdx === i) return 3;
  return 2;
}

function makeRadarOptions(
  title: string,
  emitHover: (idx: number | null) => void
): object {
  return {
    responsive: true,
    maintainAspectRatio: false,
    // Disable animation entirely — re-running the tween on every hover
    // state change is the main source of flicker when the cursor crosses
    // the tightly-packed Outputs polygons.
    animation: false,
    // Fire hover events when the cursor is anywhere near a dataset point —
    // default `intersect: true` requires pixel-perfect hits on a point marker.
    interaction: { mode: "nearest", intersect: false, axis: "xy" },
    plugins: {
      title: { display: true, text: title, font: { size: 12 } },
      legend: { display: false },
      tooltip: {
        enabled: true,
        callbacks: {
          label: (ctx: { dataset: { label?: string }; formattedValue: string }) =>
            `${ctx.dataset.label}: ${ctx.formattedValue}%`,
        },
      },
    },
    scales: {
      r: {
        min: 0,
        max: 100,
        ticks: { display: false, stepSize: 20 },
        pointLabels: { font: { size: 9 } },
      },
    },
    onHover: (_evt: unknown, elements: Array<{ datasetIndex: number }>) => {
      if (elements && elements.length > 0) {
        emitHover(elements[0].datasetIndex);
      } else {
        emitHover(null);
      }
    },
  };
}

export const DualRadarCharts: React.FC<DualRadarChartsProps> = ({
  solutions,
  inputVariables,
  outcomeNames,
  hoveredIdx,
  onHoverSolution,
}) => {
  const styles = useStyles();

  // --- Normalize inputs using InputVariable min/max bounds ---
  const inputLabels = useMemo(
    () => inputVariables.map((v) => v.name),
    [inputVariables]
  );

  // --- Normalize outputs using a *padded* observed min/max ---
  // Mapping the raw observed range directly to 0–100% makes a tight cluster
  // of solutions look spread out across the full radial axis, which is
  // visually misleading. We expand the observed range by PADDING_FACTOR
  // centered on its midpoint, so the cluster always renders as a "band
  // around the middle": with factor=3, the cluster occupies the middle 33%
  // of the radial axis (33%–67%). Tight clusters → tight band. Wider
  // observed range → wider band, but still bounded.
  const OUTPUT_PADDING_FACTOR = 3;
  const outputBounds = useMemo(() => {
    const n = outcomeNames.length;
    const mins = new Array<number>(n).fill(Number.POSITIVE_INFINITY);
    const maxs = new Array<number>(n).fill(Number.NEGATIVE_INFINITY);
    for (const sol of solutions) {
      for (let i = 0; i < n; i++) {
        const v = sol.output?.[i];
        if (typeof v !== "number" || !isFinite(v)) continue;
        if (v < mins[i]) mins[i] = v;
        if (v > maxs[i]) maxs[i] = v;
      }
    }
    return mins.map((min, i) => {
      const max = maxs[i];
      if (!isFinite(min) || !isFinite(max) || min === max) {
        // Degenerate axis — everyone plots at 50%.
        return { min: 0, max: 1, flat: true };
      }
      const center = (min + max) / 2;
      const halfRange = (max - min) / 2;
      return {
        min: center - halfRange * OUTPUT_PADDING_FACTOR,
        max: center + halfRange * OUTPUT_PADDING_FACTOR,
        flat: false,
      };
    });
  }, [solutions, outcomeNames.length]);

  const inputDatasets = useMemo(() => {
    return solutions.map((sol, i) => {
      const color = SOLUTION_COLORS[sol.runIndex % SOLUTION_COLORS.length];
      const data = inputVariables.map((v, idx) => {
        const range = v.max - v.min;
        if (range === 0) return 50;
        const raw = sol.input?.[idx] ?? v.min;
        return Math.max(0, Math.min(1, (raw - v.min) / range)) * 100;
      });
      return {
        label: `Solution ${i + 1}`,
        data,
        backgroundColor: hexToRgba(color, getFillAlpha(hoveredIdx, i)),
        borderColor: hexToRgba(color, getBorderAlpha(hoveredIdx, i)),
        borderWidth: getBorderWidth(hoveredIdx, i),
        pointRadius: hoveredIdx === i ? 3 : 2,
        pointBackgroundColor: color,
        pointBorderColor: color,
        fill: true,
      };
    });
  }, [solutions, inputVariables, hoveredIdx]);

  const outputDatasets = useMemo(() => {
    return solutions.map((sol, i) => {
      const color = SOLUTION_COLORS[sol.runIndex % SOLUTION_COLORS.length];
      const data = outcomeNames.map((_, idx) => {
        const b = outputBounds[idx];
        if (b.flat) return 50;
        const raw = sol.output?.[idx];
        if (typeof raw !== "number" || !isFinite(raw)) return 0;
        return Math.max(0, Math.min(1, (raw - b.min) / (b.max - b.min))) * 100;
      });
      return {
        label: `Solution ${i + 1}`,
        data,
        backgroundColor: hexToRgba(color, getFillAlpha(hoveredIdx, i)),
        borderColor: hexToRgba(color, getBorderAlpha(hoveredIdx, i)),
        borderWidth: getBorderWidth(hoveredIdx, i),
        pointRadius: hoveredIdx === i ? 3 : 2,
        pointBackgroundColor: color,
        pointBorderColor: color,
        fill: true,
      };
    });
  }, [solutions, outcomeNames, outputBounds, hoveredIdx]);

  // Dedupe hover updates — Chart.js onHover fires on every mouse move. Without
  // this, each move triggers a state change → full re-render → new dataset
  // objects → Chart.js re-renders → onHover fires again, producing visible
  // flicker (especially on the Outputs chart with tighter polygons).
  const lastHoverRef = useRef<number | null>(null);
  const emitHover = useCallback(
    (idx: number | null) => {
      if (lastHoverRef.current === idx) return;
      lastHoverRef.current = idx;
      onHoverSolution(idx);
    },
    [onHoverSolution]
  );

  const inputOptions = useMemo(
    () => makeRadarOptions("Input Values (normalized to min/max bounds)", emitHover),
    [emitHover]
  );
  const outputOptions = useMemo(
    () => makeRadarOptions("Output Values (centered band — padded observed range)", emitHover),
    [emitHover]
  );

  if (solutions.length === 0) return null;

  return (
    <div className={styles.wrapper}>
      {/* Legend — also drives highlight state */}
      <div
        className={styles.legendRow}
        onMouseLeave={() => onHoverSolution(null)}
      >
        {solutions.map((sol, i) => {
          const color = SOLUTION_COLORS[sol.runIndex % SOLUTION_COLORS.length];
          const opacity =
            hoveredIdx === null || hoveredIdx === i ? 1 : 0.4;
          const bg =
            hoveredIdx === i
              ? hexToRgba(color, 0.18)
              : "transparent";
          return (
            <div
              key={i}
              className={styles.legendChip}
              style={{ opacity, backgroundColor: bg }}
              onMouseEnter={() => onHoverSolution(i)}
            >
              <div
                className={styles.legendSwatch}
                style={{ backgroundColor: color }}
              />
              <Text size={200} weight={hoveredIdx === i ? "semibold" : "regular"}>
                Solution {i + 1}
                {sol.satisfied ? " ✓" : ""}
              </Text>
            </div>
          );
        })}
      </div>

      <div className={styles.chartRow}>
        <div
          className={styles.chartCard}
          onMouseLeave={() => onHoverSolution(null)}
        >
          <Text className={styles.chartTitle} weight="semibold" size={200}>
            Inputs
          </Text>
          <div className={styles.chartBox}>
            <Radar
              data={{ labels: inputLabels, datasets: inputDatasets }}
              options={inputOptions}
            />
          </div>
        </div>

        <div
          className={styles.chartCard}
          onMouseLeave={() => onHoverSolution(null)}
        >
          <Text className={styles.chartTitle} weight="semibold" size={200}>
            Outputs
          </Text>
          <div className={styles.chartBox}>
            <Radar
              data={{ labels: outcomeNames, datasets: outputDatasets }}
              options={outputOptions}
            />
          </div>
        </div>
      </div>
    </div>
  );
};
