import * as React from 'react';
import { useCallback } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  IDataSourceConfig,
  IDisplayConfig,
  IBehaviorConfig,
  IPersonalizationConfig
} from '../../../models/Configuration';
import { QuoteItem, IFieldMapping } from '../../../models/ListItem';
import { useListData } from '../../../hooks/useListData';
import { TelemetryService } from '../../../services/TelemetryService';
import { Quotes, QuoteDisplayMode } from '../../../components/quotes/Quotes';

export interface IQuotesContainerProps {
  context: WebPartContext;
  dataSource: IDataSourceConfig;
  display: IDisplayConfig;
  behavior: IBehaviorConfig;
  personalization: IPersonalizationConfig;
  mode: QuoteDisplayMode;
}

function mapQuoteItem(raw: any, _mapping: IFieldMapping): QuoteItem {
  return {
    id: raw.id || 0,
    title: raw.title || '',
    quote: raw.quote || raw.title || '',
    author: raw.author,
    role: raw.role,
    department: raw.department,
    imageUrl: raw.imageUrl || raw.image,
    category: raw.category,
    displayOrder: raw.displayOrder,
    isActive: raw.isActive !== false,
    publishDate: raw.publishDate,
    expiryDate: raw.expiryDate,
    audience: raw.audience
  };
}

export const QuotesContainer: React.FC<IQuotesContainerProps> = (props) => {
  const { context, dataSource, display, behavior, personalization, mode } = props;

  const { items, state, errorMessage, refresh } = useListData<QuoteItem>(
    context,
    dataSource,
    'Quotes',
    mapQuoteItem,
    { enableAudience: personalization.enableAudienceTargeting }
  );

  const telemetry = React.useMemo(
    () => new TelemetryService(context, 'Quotes'),
    [context]
  );

  const onItemViewed = useCallback(
    (item: QuoteItem) => {
      telemetry.track('ItemViewed', { author: item.author }, String(item.id));
    },
    [telemetry]
  );

  return (
    <Quotes
      items={items}
      state={state}
      errorMessage={errorMessage}
      display={display}
      behavior={behavior}
      mode={mode}
      onRetry={refresh}
      onItemViewed={onItemViewed}
    />
  );
};

export default QuotesContainer;
