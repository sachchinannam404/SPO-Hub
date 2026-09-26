import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Text } from '@fluentui/react/lib/Text';
import {
  IDataSourceConfig,
  IDisplayConfig,
  IBehaviorConfig,
  IPersonalizationConfig
} from '../../../models/Configuration';
import { AccordionItem, IFieldMapping } from '../../../models/ListItem';
import { useListData } from '../../../hooks/useListData';
import { AccordionRenderer } from '../../../components/accordion/Accordion';

export interface IAccordionContainerProps {
  context: WebPartContext;
  dataSource: IDataSourceConfig;
  display: IDisplayConfig;
  behavior: IBehaviorConfig;
  personalization: IPersonalizationConfig;
}

function mapAccordionItem(raw: any, _mapping: IFieldMapping): AccordionItem {
  return {
    id: raw.id || 0,
    title: raw.title || '',
    description: raw.description,
    content: raw.content,
    category: raw.category,
    link: raw.url || raw.link,
    lastUpdated: raw.lastUpdated,
    displayOrder: raw.displayOrder,
    isActive: raw.isActive !== false,
    audience: raw.audience
  };
}

export const AccordionContainer: React.FC<IAccordionContainerProps> = (props) => {
  const { context, dataSource, display, behavior, personalization } = props;

  const { items, state, errorMessage, refresh } = useListData<AccordionItem>(
    context,
    dataSource,
    'Accordion',
    mapAccordionItem,
    { enableAudience: personalization.enableAudienceTargeting }
  );

  return (
    <div>
      {display.title && (
        <Text as="h2" variant="xLarge" styles={{ root: { marginBottom: 16, fontWeight: 600 } }}>
          {display.title}
        </Text>
      )}
      <AccordionRenderer
        items={items}
        state={state}
        errorMessage={errorMessage}
        display={display}
        behavior={behavior}
        onRetry={refresh}
      />
    </div>
  );
};

export default AccordionContainer;
