import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  IDataSourceConfig,
  IDisplayConfig,
  IBehaviorConfig,
  IPersonalizationConfig
} from '../../../models/Configuration';
import { TimelineItem, IFieldMapping } from '../../../models/ListItem';
import { useListData } from '../../../hooks/useListData';
import { Timeline } from '../../../components/timeline/Timeline';

export interface ITimelineContainerProps {
  context: WebPartContext;
  dataSource: IDataSourceConfig;
  display: IDisplayConfig;
  behavior: IBehaviorConfig;
  personalization: IPersonalizationConfig;
  compact?: boolean;
}

function mapTimelineItem(raw: any, mapping: IFieldMapping): TimelineItem {
  return {
    id: raw.id || 0,
    title: raw.title || '',
    description: raw.description,
    startDate: raw.startDate,
    endDate: raw.endDate,
    status: raw.status,
    category: raw.category,
    icon: raw.icon,
    imageUrl: raw.imageUrl,
    link: raw.url || raw.link,
    owner: raw.owner,
    sequence: raw.sequence != null ? Number(raw.sequence) : undefined,
    isActive: raw.isActive !== false,
    audience: raw.audience,
    displayOrder: raw.displayOrder
  };
}

export const TimelineContainer: React.FC<ITimelineContainerProps> = (props) => {
  const { context, dataSource, display, behavior, personalization, compact } = props;

  const { items, state, errorMessage, refresh } = useListData<TimelineItem>(
    context,
    dataSource,
    'Timeline',
    mapTimelineItem,
    { enableAudience: personalization.enableAudienceTargeting }
  );

  return (
    <Timeline
      items={items}
      state={state}
      errorMessage={errorMessage}
      display={display}
      behavior={behavior}
      onRetry={refresh}
      compact={compact}
    />
  );
};

export default TimelineContainer;
