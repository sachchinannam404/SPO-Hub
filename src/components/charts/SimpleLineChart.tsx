import * as React from 'react';
import { ChartDataPoint } from '../../models/ListItem';
import styles from './Charts.module.scss';

export interface ISimpleLineChartProps {
  data: ChartDataPoint[];
  colors: string[];
  area?: boolean;
  showLegend?: boolean;
}

const WIDTH = 560;
const HEIGHT = 280;
const PAD = { top: 20, right: 20, bottom: 48, left: 48 };

export const SimpleLineChart: React.FC<ISimpleLineChartProps> = ({
  data,
  colors,
  area = false,
  showLegend = true
}) => {
  const maxVal = Math.max(...data.map(d => d.value), 1);
  const plotW = WIDTH - PAD.left - PAD.right;
  const plotH = HEIGHT - PAD.top - PAD.bottom;
  const color = colors[0] || '#0078d4';

  const points = data.map((d, i) => {
    const x = PAD.left + (data.length === 1 ? plotW / 2 : (i / (data.length - 1)) * plotW);
    const y = HEIGHT - PAD.bottom - (d.value / maxVal) * plotH;
    return { x, y, ...d };
  });

  const linePath = points.map((p, i) => `${i === 0 ? 'M' : 'L'} ${p.x} ${p.y}`).join(' ');
  const areaPath = points.length
    ? `${linePath} L ${points[points.length - 1].x} ${HEIGHT - PAD.bottom} L ${points[0].x} ${HEIGHT - PAD.bottom} Z`
    : '';

  return (
    <div className={styles.chartWrap}>
      <svg
        viewBox={`0 0 ${WIDTH} ${HEIGHT}`}
        className={styles.svg}
        role="img"
        aria-label={area ? 'Area chart' : 'Line chart'}
      >
        <line x1={PAD.left} y1={PAD.top} x2={PAD.left} y2={HEIGHT - PAD.bottom} stroke="#c8c6c4" />
        <line x1={PAD.left} y1={HEIGHT - PAD.bottom} x2={WIDTH - PAD.right} y2={HEIGHT - PAD.bottom} stroke="#c8c6c4" />
        {area && areaPath && (
          <path d={areaPath} fill={color} opacity={0.2} />
        )}
        <path d={linePath} fill="none" stroke={color} strokeWidth={2.5} />
        {points.map((p, i) => (
          <g key={i}>
            <circle cx={p.x} cy={p.y} r={4} fill={color}>
              <title>{`${p.category}: ${p.value}`}</title>
            </circle>
            <text x={p.x} y={HEIGHT - PAD.bottom + 14} textAnchor="middle" className={styles.axisLabel}>
              {p.category.length > 8 ? p.category.slice(0, 7) + '\u2026' : p.category}
            </text>
          </g>
        ))}
      </svg>
      {showLegend && (
        <ul className={styles.legend} aria-hidden="true">
          <li>
            <span className={styles.swatch} style={{ background: color }} />
            {data[0]?.series || 'Value'}
          </li>
        </ul>
      )}
    </div>
  );
};

export default SimpleLineChart;
