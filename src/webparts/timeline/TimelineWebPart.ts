import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import {
  IDataSourceConfig,
  IDisplayConfig,
  IBehaviorConfig,
  IPersonalizationConfig,
  DEFAULT_DATA_SOURCE,
  DEFAULT_DISPLAY,
  DEFAULT_BEHAVIOR,
  DEFAULT_PERSONALIZATION
} from '../../models/Configuration';
import { TimelineItem, IFieldMapping } from '../../models/ListItem';
import { TimelineContainer } from './components/TimelineContainer';

export interface ITimelineWebPartProps {
  dataSource: IDataSourceConfig;
  display: IDisplayConfig;
  behavior: IBehaviorConfig;
  personalization: IPersonalizationConfig;
  compact: boolean;
}

export default class TimelineWebPart extends BaseClientSideWebPart<ITimelineWebPartProps> {
  public render(): void {
    const element = React.createElement(TimelineContainer, {
      context: this.context,
      dataSource: this.properties.dataSource || { ...DEFAULT_DATA_SOURCE },
      display: this.properties.display || { ...DEFAULT_DISPLAY },
      behavior: this.properties.behavior || { ...DEFAULT_BEHAVIOR },
      personalization: this.properties.personalization || { ...DEFAULT_PERSONALIZATION },
      compact: !!this.properties.compact
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onInit(): Promise<void> {
    if (!this.properties.dataSource) {
      this.properties.dataSource = {
        ...DEFAULT_DATA_SOURCE,
        fieldMapping: {
          titleField: 'Title',
          descriptionField: 'Description',
          startDateField: 'StartDate',
          endDateField: 'EndDate',
          statusTimelineField: 'Status',
          categoryField: 'Category',
          iconField: 'Icon',
          ownerField: 'Owner',
          sequenceField: 'Sequence',
          linkField: 'Link',
          isActiveField: 'IsActive',
          audienceField: 'Audience'
        } as IFieldMapping
      };
    }
    if (!this.properties.display) this.properties.display = { ...DEFAULT_DISPLAY };
    if (!this.properties.behavior) this.properties.behavior = { ...DEFAULT_BEHAVIOR };
    if (!this.properties.personalization) {
      this.properties.personalization = { ...DEFAULT_PERSONALIZATION };
    }
    return super.onInit();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Timeline Data Source' },
          groups: [
            {
              groupName: 'SharePoint List',
              groupFields: [
                PropertyPaneTextField('dataSource.listTitle', {
                  label: 'List Title'
                }),
                PropertyPaneTextField('dataSource.orderBy', {
                  label: 'Order By Field',
                  description: 'Defaults to Sequence'
                }),
                PropertyPaneSlider('dataSource.itemLimit', {
                  label: 'Item Limit',
                  min: 1,
                  max: 100,
                  step: 1,
                  showValue: true
                }),
                PropertyPaneToggle('compact', {
                  label: 'Compact mode'
                }),
                PropertyPaneToggle('personalization.enableAudienceTargeting', {
                  label: 'Enable Audience Targeting'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
