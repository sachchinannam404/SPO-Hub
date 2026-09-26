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
import { IFieldMapping } from '../../models/ListItem';
import { AccordionContainer } from './components/AccordionContainer';

export interface IAccordionWebPartProps {
  dataSource: IDataSourceConfig;
  display: IDisplayConfig;
  behavior: IBehaviorConfig;
  personalization: IPersonalizationConfig;
}

export default class AccordionWebPart extends BaseClientSideWebPart<IAccordionWebPartProps> {
  public render(): void {
    const element = React.createElement(AccordionContainer, {
      context: this.context,
      dataSource: this.properties.dataSource || { ...DEFAULT_DATA_SOURCE },
      display: this.properties.display || { ...DEFAULT_DISPLAY },
      behavior: this.properties.behavior || { ...DEFAULT_BEHAVIOR },
      personalization: this.properties.personalization || { ...DEFAULT_PERSONALIZATION }
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
          contentField: 'Content',
          categoryField: 'Category',
          orderField: 'DisplayOrder',
          isActiveField: 'IsActive',
          audienceField: 'Audience',
          linkField: 'Link'
        } as IFieldMapping
      };
    }
    if (!this.properties.display) this.properties.display = { ...DEFAULT_DISPLAY };
    if (!this.properties.behavior) {
      this.properties.behavior = {
        ...DEFAULT_BEHAVIOR,
        enableSearch: true,
        allowMultipleExpanded: false,
        expandFirstItem: true
      };
    }
    if (!this.properties.personalization) {
      this.properties.personalization = { ...DEFAULT_PERSONALIZATION };
    }
    return super.onInit();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Accordion / FAQ Data Source' },
          groups: [
            {
              groupName: 'SharePoint List',
              groupFields: [
                PropertyPaneTextField('dataSource.listTitle', { label: 'List Title' }),
                PropertyPaneTextField('dataSource.filter', { label: 'OData Filter (optional)' }),
                PropertyPaneTextField('dataSource.orderBy', { label: 'Order By Field' }),
                PropertyPaneSlider('dataSource.itemLimit', {
                  label: 'Item Limit', min: 1, max: 100, step: 1, showValue: true
                })
              ]
            },
            {
              groupName: 'Field Mapping',
              groupFields: [
                PropertyPaneTextField('dataSource.fieldMapping.titleField', { label: 'Title Field' }),
                PropertyPaneTextField('dataSource.fieldMapping.contentField', { label: 'Content Field' }),
                PropertyPaneTextField('dataSource.fieldMapping.categoryField', { label: 'Category Field' }),
                PropertyPaneTextField('dataSource.fieldMapping.orderField', { label: 'Order Field' })
              ]
            },
            {
              groupName: 'Behavior',
              groupFields: [
                PropertyPaneToggle('behavior.enableSearch', { label: 'Enable Search' }),
                PropertyPaneToggle('behavior.allowMultipleExpanded', { label: 'Allow Multiple Expanded' }),
                PropertyPaneToggle('behavior.expandFirstItem', { label: 'Expand First Item' }),
                PropertyPaneToggle('personalization.enableAudienceTargeting', { label: 'Enable Audience Targeting' })
              ]
            }
          ]
        }
      ]
    };
  }
}
