import * as React from 'react';
import { ChartDataPoint } from '../../models/ListItem';
import styles from './Charts.module.scss';

export interface ISimplePieChartProps {
  data: ChartDataPoint[];
  colors: string[];
  doughnut?: boolean;
  showLegend?: boolean;
}

const SIZE = 280;
const CX = SIZE / 2;
const CY = SIZE / 2;
const R = 100;

export const SimplePieChart: React.FC<ISimplePieChartProps> = ({
  data,
  colors,
  doughnut = false,
  showLegend = true
}) => {
  const total = data.reduce((s, d) => s + Math.max(d.value, 0), 0) || 1;
  let angle = -Math.PI / 2;

  const slices = data.map((d, i) => {
    const sliceAngle = (Math.max(d.value, 0) / total) * Math.PI * 2;
    const start = angle;
    const end = angle + sliceAngle;
    angle = end;
    const large = sliceAngle > Math.PI ? 1 : 0;
    const x1 = CX + R * Math.cos(start);
    const y1 = CY + R * Math.sin(start);
    const x2 = CX + R * Math.cos(end);
    const y2 = CY + R * Math.sin(end);
    const path = [
      `M ${CX} ${CY}`,
      `L ${x1} ${y1}`,
      `A ${R} ${R} 0 ${large} 1 ${x2} ${y2}`,
      'Z'
    ].join(' ');
    return {
      path,
      color: d.color || colors[i % colors.length],
      label: d.category,
      value: d.value,
      pct: ((d.value / total) * 100).toFixed(1)
    };
  });

  return (
    <div className={styles.chartWrap}>
      <svg
        viewBox={`0 0 ${SIZE} ${SIZE}`}
        className={styles.svgPie}
        role="img"
        aria-label={doughnut ? 'Doughnut chart' : 'Pie chart'}
      >
        {slices.map((s, i) => (
          <path key={i} d={s.path} fill={s.color}>
            <title>{`${s.label}: ${s.value} (${s.pct}%)`}</title>
          </path>
        ))}
        {doughnut && (
          <circle cx={CX} cy={CY} r={R * 0.55} fill="#ffffff" />
        )}
      </svg>
      {showLegend && (
        <ul className={styles.legend}>
          {slices.map((s, i) => (
            <li key={i}>
              <span className={styles.swatch} style={{ background: s.color }} />
              {s.label} ({s.pct}%)
            </li>
          ))}
        </ul>
      )}
    </div>
  );
};

export default SimplePieChart;
