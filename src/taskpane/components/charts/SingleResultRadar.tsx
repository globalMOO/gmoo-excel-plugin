// Radar plots for a single (non-Multi-Solve) optimization result.
//
// Two charts, raw / non-normalized per product guidance — a single normalized
// polygon degenerates to a uniform ring and tells the user nothing:
//   • Inputs  — one polygon of the optimal input values.
//   • Outputs — two polygons, Achieved vs Target, so the user can see how close
//     each outcome landed. Minimize/Maximize outcomes have no meaningful target,
//     so their Target point is dropped (Chart.js renders a gap).
import React, { useMemo } from "react";
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
import type { Inverse } from "../../types/gmoo";
import { buildSingleResultRadarData } from "./singleResultRadarData";

ChartJS.register(
  RadialLinearScale,
  PointElement,
  LineElement,
  Filler,
  Tooltip,
  Legend,
  RadarController
);

const ACHIEVED_COLOR = "#3498db";
const TARGET_COLOR = "#e94560";
const INPUT_COLOR = "#27ae60";

const useStyles = makeStyles({
  wrapper: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
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
});

function baseOptions(title: string, showLegend: boolean): object {
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: false,
    plugins: {
      title: { display: true, text: title, font: { size: 12 } },
      legend: { display: showLegend, position: "bottom" as const },
    },
    scales: {
      r: {
        // Raw values — let Chart.js pick the radial scale; begin at zero only
        // when all values are non-negative so the polygon isn't pushed off-axis.
        ticks: { font: { size: 9 } },
        pointLabels: { font: { size: 9 } },
      },
    },
  };
}

function hexToRgba(hex: string, alpha: number): string {
  const c = hex.replace("#", "");
  const r = parseInt(c.substring(0, 2), 16);
  const g = parseInt(c.substring(2, 4), 16);
  const b = parseInt(c.substring(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

interface SingleResultRadarProps {
  inverse: Inverse;
  inputVariableNames: string[];
  outcomeNames: string[];
}

export const SingleResultRadar: React.FC<SingleResultRadarProps> = ({
  inverse,
  inputVariableNames,
  outcomeNames,
}) => {
  const styles = useStyles();
  const data = useMemo(
    () => buildSingleResultRadarData(inverse, inputVariableNames, outcomeNames),
    [inverse, inputVariableNames, outcomeNames]
  );

  const inputChartData = {
    labels: data.inputLabels,
    datasets: [
      {
        label: "Optimal input",
        data: data.inputValues,
        backgroundColor: hexToRgba(INPUT_COLOR, 0.25),
        borderColor: hexToRgba(INPUT_COLOR, 0.9),
        borderWidth: 2,
        pointRadius: 3,
        pointBackgroundColor: INPUT_COLOR,
        fill: true,
      },
    ],
  };

  const outputChartData = {
    labels: data.outputLabels,
    datasets: [
      {
        label: "Achieved",
        data: data.achieved,
        backgroundColor: hexToRgba(ACHIEVED_COLOR, 0.25),
        borderColor: hexToRgba(ACHIEVED_COLOR, 0.9),
        borderWidth: 2,
        pointRadius: 3,
        pointBackgroundColor: ACHIEVED_COLOR,
        fill: true,
      },
      ...(data.hasAnyTarget
        ? [
            {
              label: "Target",
              data: data.target,
              backgroundColor: hexToRgba(TARGET_COLOR, 0.1),
              borderColor: hexToRgba(TARGET_COLOR, 0.9),
              borderWidth: 2,
              borderDash: [5, 4],
              pointRadius: 3,
              pointBackgroundColor: TARGET_COLOR,
              fill: false,
              spanGaps: false,
            },
          ]
        : []),
    ],
  };

  return (
    <div className={styles.wrapper}>
      {data.inputLabels.length >= 3 && (
        <div className={styles.chartCard}>
          <Text className={styles.chartTitle} weight="semibold" size={200}>
            Optimal Inputs
          </Text>
          <div className={styles.chartBox}>
            <Radar data={inputChartData} options={baseOptions("Inputs (raw values)", false)} />
          </div>
        </div>
      )}
      {data.outputLabels.length >= 3 && (
        <div className={styles.chartCard}>
          <Text className={styles.chartTitle} weight="semibold" size={200}>
            Outcomes — Achieved vs Target
          </Text>
          <div className={styles.chartBox}>
            <Radar
              data={outputChartData}
              options={baseOptions("Outcomes (raw values)", data.hasAnyTarget)}
            />
          </div>
        </div>
      )}
    </div>
  );
};
