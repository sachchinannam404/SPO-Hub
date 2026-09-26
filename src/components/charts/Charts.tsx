/**
 * List-Based Charts Renderer
 * Generic analytics renderer driven by SharePoint list data.
 * Lightweight SVG implementation (no external chart library).
 */

import * as React from 'react';
import { useMemo, useState } from 'react';
import { Text } from '@fluentui/react/lib/Text';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import {
  ChartDataPoint,
  ChartType,
  ChartSeriesConfig
} from '../../models/ListItem';
import { IDisplayConfig, IChartConfig } from '../../models/Configuration';
import { LoadingState } from '../common/LoadingState/LoadingState';
import { EmptyState } from '../common/EmptyState/EmptyState';
import { ErrorState } from '../common/ErrorState/ErrorState';
import { UnauthorizedState } from '../common/UnauthorizedState/UnauthorizedState';
import { LoadState } from '../../hooks/useListData';
import { SimpleBarChart } from './SimpleBarChart';
import { SimplePieChart } from './SimplePieChart';
import { SimpleLineChart } from './SimpleLineChart';
import styles from './Charts.module.scss';

export interface IChartsProps {
  rawItems: any[];
  state: LoadState;
  errorMessage?: string;
  display: IDisplayConfig;
  chartConfig: IChartConfig;
  onRetry?: () => void;
}

const DEFAULT_COLORS = [
  '#0078d4', '#107c10', '#ffaa44', '#d13438',
  '#8764b8', '#00b7c3', '#ca5010', '#69797e'
];

export const Charts: React.FC<IChartsProps> = ({
  rawItems,
  state,
  errorMessage,
  display,
  chartConfig,
  onRetry
}) => {
  const [chartTypeOverride, setChartTypeOverride] = useState<ChartType | null>(null);
  const chartType = chartTypeOverride || chartConfig.chartType || 'bar';

  const dataPoints = useMemo(() => {
    return aggregateItems(rawItems, chartConfig);
  }, [rawItems, chartConfig]);

  const typeOptions: IDropdownOption[] = [
    { key: 'bar', text: 'Bar' },
    { key: 'horizontalBar', text: 'Horizontal Bar' },
    { key: 'line', text: 'Line' },
    { key: 'area', text: 'Area' },
    { key: 'pie', text: 'Pie' },
    { key: 'doughnut', text: 'Doughnut' }
  ];

  if (state === 'loading' || state === 'idle') {
    return <LoadingState label="Loading chart data..." />;
  }
  if (state === 'unauthorized') return <UnauthorizedState />;
  if (state === 'configError' || state === 'error') {
    return (
      <ErrorState
        message={errorMessage}
        onRetry={onRetry}
        isConfigError={state === 'configError'}
      />
    );
  }
  if (state === 'empty' || dataPoints.length === 0) {
    return <EmptyState message="No chart data is currently available." />;
  }

  const colors = chartConfig.colors && chartConfig.colors.length > 0
    ? chartConfig.colors
    : DEFAULT_COLORS;

  const title = chartConfig.title || display.title;

  return (
    <div className={styles.root} role="region" aria-label={title || 'Chart'}>
      <div className={styles.header}>
        {title && (
          <Text as="h2" variant="xLarge" className={styles.title}>
            {title}
          </Text>
        )}
        <Dropdown
          options={typeOptions}
          selectedKey={chartType}
          onChange={(_, opt) => setChartTypeOverride((opt?.key as ChartType) || null)}
          styles={{ root: { minWidth: 140 } }}
          ariaLabel="Chart type"
        />
      </div>

      <div className={styles.chartArea}>
        {(chartType === 'bar' || chartType === 'horizontalBar') && (
          <SimpleBarChart
            data={dataPoints}
            colors={colors}
            horizontal={chartType === 'horizontalBar'}
            showLegend={chartConfig.showLegend !== false}
          />
        )}
        {(chartType === 'line' || chartType === 'area') && (
          <SimpleLineChart
            data={dataPoints}
            colors={colors}
            area={chartType === 'area'}
            showLegend={chartConfig.showLegend !== false}
          />
        )}
        {(chartType === 'pie' || chartType === 'doughnut') && (
          <SimplePieChart
            data={dataPoints}
            colors={colors}
            doughnut={chartType === 'doughnut'}
            showLegend={chartConfig.showLegend !== false}
          />
        )}
      </div>

      <table className={styles.srOnly}>
        <caption>{title || 'Chart data'}</caption>
        <thead>
          <tr>
            <th scope="col">Category</th>
            <th scope="col">Value</th>
          </tr>
        </thead>
        <tbody>
          {dataPoints.map((dp, i) => (
            <tr key={i}>
              <td>{dp.category}</td>
              <td>{dp.value}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

function aggregateItems(items: any[], config: IChartConfig): ChartDataPoint[] {
  const categoryField = config.categoryField || 'Title';
  const series = config.series && config.series.length > 0
    ? config.series
    : [{ field: 'Value', aggregation: 'sum' as const }];

  const groups = new Map<string, any[]>();
  items.forEach(item => {
    const cat = String(item[categoryField] ?? item.category ?? item.title ?? 'Unknown');
    if (!groups.has(cat)) groups.set(cat, []);
    groups.get(cat)!.push(item);
  });

  const primary = series[0];
  const points: ChartDataPoint[] = [];

  groups.forEach((groupItems, category) => {
    const value = aggregate(groupItems, primary);
    points.push({
      category,
      value,
      series: primary.label || primary.field,
      color: primary.color
    });
  });

  return points.sort((a, b) => a.category.localeCompare(b.category));
}

function aggregate(items: any[], series: ChartSeriesConfig): number {
  const field = series.field;
  const values = items.map(i => {
    const v = i[field] ?? i.value ?? i[field.toLowerCase()];
    const n = typeof v === 'number' ? v : parseFloat(v);
    return isNaN(n) ? 0 : n;
  });

  switch (series.aggregation) {
    case 'count':
      return items.length;
    case 'avg':
      return values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0;
    case 'min':
      return values.length ? Math.min(...values) : 0;
    case 'max':
      return values.length ? Math.max(...values) : 0;
    case 'sum':
    default:
      return values.reduce((a, b) => a + b, 0);
  }
}

export default Charts;
