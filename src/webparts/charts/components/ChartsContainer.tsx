import * as React from 'react';
import { useMemo } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  IDataSourceConfig,
  IDisplayConfig,
  IChartConfig
} from '../../../models/Configuration';
import { IFieldMapping } from '../../../models/ListItem';
import { useListData } from '../../../hooks/useListData';
import { Charts } from '../../../components/charts/Charts';

export interface IChartsContainerProps {
  context: WebPartContext;
  dataSource: IDataSourceConfig;
  display: IDisplayConfig;
  chartConfig: IChartConfig;
}

/** Pass-through mapper: charts need raw field access for aggregation */
function mapChartRow(raw: any, _mapping: IFieldMapping): any {
  return { ...raw, id: raw.id || 0, title: raw.title || raw.Title || '' };
}

export const ChartsContainer: React.FC<IChartsContainerProps> = (props) => {
  const { context, dataSource, display, chartConfig } = props;

  const enrichedDataSource = useMemo(() => {
    const mapping = { ...(dataSource.fieldMapping || {}) };
    if (chartConfig.categoryField && !mapping.categoryField) {
      mapping.categoryField = chartConfig.categoryField;
    }
    if (chartConfig.series?.[0]?.field) {
      (mapping as any).valueField = chartConfig.series[0].field;
    }
    return { ...dataSource, fieldMapping: mapping };
  }, [dataSource, chartConfig]);

  const { items, state, errorMessage, refresh } = useListData<any>(
    context,
    enrichedDataSource,
    'Charts',
    mapChartRow,
    {
      extraSelect: buildExtraSelect(chartConfig)
    }
  );

  return (
    <Charts
      rawItems={items}
      state={state}
      errorMessage={errorMessage}
      display={display}
      chartConfig={chartConfig}
      onRetry={refresh}
    />
  );
};

function buildExtraSelect(config: IChartConfig): string[] {
  const fields: string[] = [];
  if (config.categoryField) fields.push(config.categoryField);
  (config.series || []).forEach(s => {
    if (s.field) fields.push(s.field);
  });
  return fields;
}

export default ChartsContainer;
