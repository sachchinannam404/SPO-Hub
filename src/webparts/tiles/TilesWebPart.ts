import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneDropdown
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
import { TileItem, IFieldMapping } from '../../models/ListItem';
import { TilesContainer } from './components/TilesContainer';

export interface ITilesWebPartProps {
  dataSource: IDataSourceConfig;
  display: IDisplayConfig;
  behavior: IBehaviorConfig;
  personalization: IPersonalizationConfig;
}

export default class TilesWebPart extends BaseClientSideWebPart<ITilesWebPartProps> {
  public render(): void {
    const element = React.createElement(TilesContainer, {
      context: this.context,
      dataSource: this.properties.dataSource || { ...DEFAULT_DATA_SOURCE },
      display: this.properties.display || { ...DEFAULT_DISPLAY },
      behavior: this.properties.behavior || { ...DEFAULT_BEHAVIOR },
      personalization: this.properties.personalization || { ...DEFAULT_PERSONALIZATION },
      webPartWidth: this.domElement.clientWidth
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
    // Ensure defaults
    if (!this.properties.dataSource) {
      this.properties.dataSource = {
        ...DEFAULT_DATA_SOURCE,
        fieldMapping: {
          titleField: 'Title',
          descriptionField: 'Description',
          imageField: 'Image',
          linkField: 'Link',
          categoryField: 'Category',
          orderField: 'DisplayOrder',
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
          header: { description: 'Data Source' },
          groups: [
            {
              groupName: 'SharePoint List',
              groupFields: [
                PropertyPaneTextField('dataSource.listTitle', {
                  label: 'List Title',
                  description: 'Title of the SharePoint list that holds the tiles'
                }),
                PropertyPaneTextField('dataSource.listId', {
                  label: 'List ID (optional)',
                  description: 'GUID of the list – takes precedence over title'
                }),
                PropertyPaneTextField('dataSource.filter', {
                  label: 'OData Filter (optional)'
                }),
                PropertyPaneTextField('dataSource.orderBy', {
                  label: 'Order By Field'
                }),
                PropertyPaneDropdown('dataSource.orderDirection', {
                  label: 'Order Direction',
                  options: [
                    { key: 'asc', text: 'Ascending' },
                    { key: 'desc', text: 'Descending' }
                  ]
                }),
                PropertyPaneSlider('dataSource.itemLimit', {
                  label: 'Item Limit',
                  min: 1,
                  max: 100,
                  step: 1,
                  showValue: true
                })
              ]
            },
            {
              groupName: 'Field Mapping',
              groupFields: [
                PropertyPaneTextField('dataSource.fieldMapping.titleField', {
                  label: 'Title Field'
                }),
                PropertyPaneTextField('dataSource.fieldMapping.descriptionField', {
                  label: 'Description Field'
                }),
                PropertyPaneTextField('dataSource.fieldMapping.imageField', {
                  label: 'Image Field'
                }),
                PropertyPaneTextField('dataSource.fieldMapping.linkField', {
                  label: 'Link Field'
                }),
                PropertyPaneTextField('dataSource.fieldMapping.categoryField', {
                  label: 'Category Field'
                }),
                PropertyPaneTextField('dataSource.fieldMapping.orderField', {
                  label: 'Order Field'
                }),
                PropertyPaneTextField('dataSource.fieldMapping.isActiveField', {
                  label: 'Is Active Field'
                })
              ]
            }
          ]
        },
        {
          header: { description: 'Display & Behavior' },
          groups: [
            {
              groupName: 'Display',
              groupFields: [
                PropertyPaneTextField('display.title', {
                  label: 'Web Part Title'
                }),
                PropertyPaneDropdown('display.layout', {
                  label: 'Layout',
                  options: [
                    { key: 'grid', text: 'Grid' },
                    { key: 'list', text: 'List' }
                  ]
                }),
                PropertyPaneSlider('display.columns', {
                  label: 'Columns (desktop)',
                  min: 1,
                  max: 6,
                  step: 1,
                  showValue: true
                }),
                PropertyPaneDropdown('display.cardSize', {
                  label: 'Card Size',
                  options: [
                    { key: 'small', text: 'Small' },
                    { key: 'medium', text: 'Medium' },
                    { key: 'large', text: 'Large' }
                  ]
                }),
                PropertyPaneToggle('display.showDescription', {
                  label: 'Show Description'
                }),
                PropertyPaneToggle('display.showCategory', {
                  label: 'Show Category'
                })
              ]
            },
            {
              groupName: 'Behavior',
              groupFields: [
                PropertyPaneToggle('behavior.enableSearch', {
                  label: 'Enable Search'
                }),
                PropertyPaneToggle('behavior.enableFilter', {
                  label: 'Enable Category Filter'
                }),
                PropertyPaneToggle('behavior.openLinksInNewTab', {
                  label: 'Open Links in New Tab'
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
