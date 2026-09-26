import * as React from 'react';
import { ChartDataPoint } from '../../models/ListItem';
import styles from './Charts.module.scss';

export interface ISimpleBarChartProps {
  data: ChartDataPoint[];
  colors: string[];
  horizontal?: boolean;
  showLegend?: boolean;
}

const WIDTH = 560;
const HEIGHT = 280;
const PAD = { top: 20, right: 20, bottom: 48, left: 48 };

export const SimpleBarChart: React.FC<ISimpleBarChartProps> = ({
  data,
  colors,
  horizontal = false,
  showLegend = true
}) => {
  const maxVal = Math.max(...data.map(d => d.value), 1);
  const plotW = WIDTH - PAD.left - PAD.right;
  const plotH = HEIGHT - PAD.top - PAD.bottom;

  if (horizontal) {
    const barH = Math.min(28, (plotH / data.length) * 0.7);
    const gap = plotH / data.length;

    return (
      <div className={styles.chartWrap}>
        <svg viewBox={`0 0 ${WIDTH} ${HEIGHT}`} className={styles.svg} role="img" aria-label="Horizontal bar chart">
          {data.map((d, i) => {
            const y = PAD.top + i * gap + (gap - barH) / 2;
            const w = (d.value / maxVal) * plotW;
            const color = d.color || colors[i % colors.length];
            return (
              <g key={i}>
                <rect x={PAD.left} y={y} width={Math.max(w, 1)} height={barH} fill={color} rx={2}>
                  <title>{`${d.category}: ${d.value}`}</title>
                </rect>
                <text x={PAD.left - 6} y={y + barH / 2} textAnchor="end" dominantBaseline="middle" className={styles.axisLabel}>
                  {truncate(d.category, 10)}
                </text>
                <text x={PAD.left + w + 4} y={y + barH / 2} dominantBaseline="middle" className={styles.valueLabel}>
                  {formatNum(d.value)}
                </text>
              </g>
            );
          })}
        </svg>
        {showLegend && <Legend data={data} colors={colors} />}
      </div>
    );
  }

  const barW = Math.min(40, (plotW / data.length) * 0.6);
  const gap = plotW / data.length;

  return (
    <div className={styles.chartWrap}>
      <svg viewBox={`0 0 ${WIDTH} ${HEIGHT}`} className={styles.svg} role="img" aria-label="Bar chart">
        <line x1={PAD.left} y1={PAD.top} x2={PAD.left} y2={HEIGHT - PAD.bottom} stroke="#c8c6c4" />
        <line x1={PAD.left} y1={HEIGHT - PAD.bottom} x2={WIDTH - PAD.right} y2={HEIGHT - PAD.bottom} stroke="#c8c6c4" />
        {data.map((d, i) => {
          const h = (d.value / maxVal) * plotH;
          const x = PAD.left + i * gap + (gap - barW) / 2;
          const y = HEIGHT - PAD.bottom - h;
          const color = d.color || colors[i % colors.length];
          return (
            <g key={i}>
              <rect x={x} y={y} width={barW} height={Math.max(h, 1)} fill={color} rx={2}>
                <title>{`${d.category}: ${d.value}`}</title>
              </rect>
              <text x={x + barW / 2} y={HEIGHT - PAD.bottom + 14} textAnchor="middle" className={styles.axisLabel}>
                {truncate(d.category, 8)}
              </text>
              <text x={x + barW / 2} y={y - 4} textAnchor="middle" className={styles.valueLabel}>
                {formatNum(d.value)}
              </text>
            </g>
          );
        })}
      </svg>
      {showLegend && <Legend data={data} colors={colors} />}
    </div>
  );
};

const Legend: React.FC<{ data: ChartDataPoint[]; colors: string[] }> = ({ data, colors }) => (
  <ul className={styles.legend} aria-hidden="true">
    {data.map((d, i) => (
      <li key={i}>
        <span className={styles.swatch} style={{ background: d.color || colors[i % colors.length] }} />
        {d.category}
      </li>
    ))}
  </ul>
);

function truncate(s: string, n: number): string {
  return s.length > n ? s.slice(0, n - 1) + '\u2026' : s;
}

function formatNum(n: number): string {
  if (Math.abs(n) >= 1000000) return (n / 1000000).toFixed(1) + 'M';
  if (Math.abs(n) >= 1000) return (n / 1000).toFixed(1) + 'k';
  return Number.isInteger(n) ? String(n) : n.toFixed(1);
}

export default SimpleBarChart;
