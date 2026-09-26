/**
 * Common Configuration Framework
 * Shared property contracts for all web parts.
 */

import { IFieldMapping, ChartType, ChartSeriesConfig } from './ListItem';

export interface IDataSourceConfig {
  listId?: string;
  listTitle?: string;
  viewId?: string;
  viewTitle?: string;
  filter?: string;
  orderBy?: string;
  orderDirection?: 'asc' | 'desc';
  itemLimit?: number;
  refreshIntervalSeconds?: number;
  fieldMapping?: IFieldMapping;
}

export interface IDisplayConfig {
  title?: string;
  description?: string;
  layout?: 'grid' | 'list' | 'carousel' | 'compact' | 'cards';
  columns?: number;
  cardSize?: 'small' | 'medium' | 'large';
  imagePosition?: 'top' | 'left' | 'right' | 'background';
  showMore?: boolean;
  showMetadata?: boolean;
  showCategory?: boolean;
  showDescription?: boolean;
}

export interface IBehaviorConfig {
  enableSearch?: boolean;
  enableFilter?: boolean;
  enablePagination?: boolean;
  enableDeepLink?: boolean;
  openLinksInNewTab?: boolean;
  autoRefresh?: boolean;
  allowMultipleExpanded?: boolean;
  expandFirstItem?: boolean;
  autoRotate?: boolean;
  autoRotateIntervalSeconds?: number;
}

export interface IPersonalizationConfig {
  enableAudienceTargeting?: boolean;
  department?: string;
  country?: string;
  role?: string;
  groupIds?: string[];
}

export interface IAutomationConfig {
  enableAction?: boolean;
  powerAutomateFlowUrl?: string;
  actionLabel?: string;
  confirmationRequired?: boolean;
  successMessage?: string;
  failureMessage?: string;
}

export interface IChartConfig {
  chartType: ChartType;
  categoryField: string;
  series: ChartSeriesConfig[];
  showLegend?: boolean;
  showTitle?: boolean;
  title?: string;
  colors?: string[];
}

/** Combined web part properties base */
export interface IListDrivenWebPartProps {
  dataSource: IDataSourceConfig;
  display: IDisplayConfig;
  behavior: IBehaviorConfig;
  personalization: IPersonalizationConfig;
  automation: IAutomationConfig;
}

export const DEFAULT_DATA_SOURCE: IDataSourceConfig = {
  itemLimit: 50,
  orderDirection: 'asc',
  refreshIntervalSeconds: 0
};

export const DEFAULT_DISPLAY: IDisplayConfig = {
  layout: 'grid',
  columns: 4,
  cardSize: 'medium',
  imagePosition: 'top',
  showMore: true,
  showMetadata: false,
  showCategory: true,
  showDescription: true
};

export const DEFAULT_BEHAVIOR: IBehaviorConfig = {
  enableSearch: false,
  enableFilter: false,
  enablePagination: false,
  enableDeepLink: false,
  openLinksInNewTab: true,
  autoRefresh: false,
  allowMultipleExpanded: false,
  expandFirstItem: true,
  autoRotate: false,
  autoRotateIntervalSeconds: 8
};

export const DEFAULT_PERSONALIZATION: IPersonalizationConfig = {
  enableAudienceTargeting: false
};

export const DEFAULT_AUTOMATION: IAutomationConfig = {
  enableAction: false,
  confirmationRequired: true
};
