import React from "react";
import { Line } from "react-chartjs-2";
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  LogarithmicScale,
  PointElement,
  LineElement,
  Title,
  Tooltip,
  Legend,
} from "chart.js";
import { filteredL1Norm, type Inverse } from "../../types/gmoo";

// Register Chart.js components
ChartJS.register(
  CategoryScale,
  LinearScale,
  LogarithmicScale,
  PointElement,
  LineElement,
  Title,
  Tooltip,
  Legend
);

interface ConvergenceChartProps {
  iterations: Inverse[];
}

export const ConvergenceChart: React.FC<ConvergenceChartProps> = ({ iterations }) => {
  if (iterations.length === 0) return null;

  const data = {
    labels: iterations.map((inv) => inv.iteration),
    datasets: [
      {
        label: "Error",
        data: iterations.map((inv) => filteredL1Norm(inv)),
        borderColor: "#0078d4",
        backgroundColor: "rgba(0, 120, 212, 0.1)",
        borderWidth: 2,
        pointRadius: 2,
        tension: 0.1,
      },
    ],
  };

  const options = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      title: {
        display: true,
        text: "Convergence",
        font: { size: 12 },
      },
      legend: {
        display: false,
      },
    },
    scales: {
      x: {
        title: {
          display: true,
          text: "Iteration",
          font: { size: 10 },
        },
        ticks: {
          font: { size: 9 },
          stepSize: 5,
        },
      },
      y: {
        type: "logarithmic" as const,
        title: {
          display: true,
          text: "Error",
          font: { size: 10 },
        },
        ticks: { font: { size: 9 } },
      },
    },
  };

  return (
    <div style={{ height: "200px", width: "100%" }}>
      <Line data={data} options={options} />
    </div>
  );
};
