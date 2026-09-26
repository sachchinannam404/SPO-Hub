import * as React from 'react';
import { useCallback } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Text } from '@fluentui/react/lib/Text';
import {
  IDataSourceConfig,
  IDisplayConfig,
  IBehaviorConfig,
  IPersonalizationConfig
} from '../../../models/Configuration';
import { TileItem, IFieldMapping } from '../../../models/ListItem';
import { useListData } from '../../../hooks/useListData';
import { TelemetryService } from '../../../services/TelemetryService';
import { Tiles } from '../../../components/tiles/Tiles';

export interface ITilesContainerProps {
  context: WebPartContext;
  dataSource: IDataSourceConfig;
  display: IDisplayConfig;
  behavior: IBehaviorConfig;
  personalization: IPersonalizationConfig;
  webPartWidth: number;
}

function mapTileItem(raw: any, mapping: IFieldMapping): TileItem {
  return {
    id: raw.id || 0,
    title: raw.title || raw.Title || '',
    description: raw.description,
    imageUrl: raw.imageUrl || raw.image,
    altText: raw.altText || raw.title,
    url: raw.url || raw.link,
    openInNewTab: true,
    category: raw.category,
    displayOrder: raw.displayOrder,
    isActive: raw.isActive !== false,
    audience: raw.audience,
    iconName: raw.iconName || raw.icon
  };
}

export const TilesContainer: React.FC<ITilesContainerProps> = (props) => {
  const { context, dataSource, display, behavior, personalization, webPartWidth } = props;

  const { items, state, errorMessage, refresh } = useListData<TileItem>(
    context,
    dataSource,
    'Tiles',
    mapTileItem,
    {
      enableAudience: personalization.enableAudienceTargeting
    }
  );

  const telemetry = React.useMemo(
    () => new TelemetryService(context, 'Tiles'),
    [context]
  );

  const onItemClick = useCallback(
    (item: TileItem) => {
      telemetry.trackItemClicked(String(item.id), {
        title: item.title,
        category: item.category
      });
      if (item.url) {
        telemetry.track('LinkOpened', { urlHost: tryGetHost(item.url) });
      }
    },
    [telemetry]
  );

  return (
    <div>
      {display.title && (
        <Text
          as="h2"
          variant="xLarge"
          styles={{ root: { marginBottom: 16, fontWeight: 600 } }}
        >
          {display.title}
        </Text>
      )}
      <Tiles
        items={items}
        state={state}
        errorMessage={errorMessage}
        display={display}
        behavior={behavior}
        onRetry={refresh}
        onItemClick={onItemClick}
        webPartWidth={webPartWidth}
      />
    </div>
  );
};

function tryGetHost(url: string): string {
  try {
    return new URL(url).host;
  } catch {
    return 'invalid';
  }
}

export default TilesContainer;
